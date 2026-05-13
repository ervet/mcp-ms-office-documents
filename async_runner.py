"""Async-safe wrapper for blocking document-generation calls.

Every Office-document generator in this project (``markdown_to_word``,
``markdown_to_excel``, ``create_presentation``, ``create_eml``,
``create_xml_file``, and every dynamically-registered template tool) is a
synchronous, blocking function. They perform:

  * file I/O (opening .docx / .pptx templates, which are zip archives)
  * CPU-bound markdown / mustache rendering
  * synchronous network I/O (``requests.get`` for image downloads,
    ``boto3`` S3 uploads)

FastMCP runs on top of an asyncio event loop (uvicorn / Starlette).
Calling a blocking function directly from an ``async def`` MCP tool
handler freezes the loop for the full duration of the call — no other
request can be served, including the Kubernetes liveness / readiness
probes. Repeated probe timeouts cause kubelet to SIGTERM the pod, which
is the failure mode observed in EKS:

    ERROR: ASGI callable returned without completing response.
    ERROR: Cancel 0 running task(s), timeout graceful shutdown exceeded

(The "0 running tasks" line is the giveaway: there were no async tasks
because the work was running synchronously on the event loop itself.)

Whether blocking work is offloaded to a worker thread is controlled by
the ``RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED`` environment variable
(see ``config.Config.run_blocking_by_asyncio_thread_enabled``):

  * Disabled (default) — call the sync function inline on the event
    loop. The loop is blocked until the call returns. This is the
    legacy behaviour for static tools; dynamic tools (previously
    registered as sync ``def`` and auto-threaded by FastMCP) will also
    run on the event loop in this mode, since the dynamic handlers are
    now ``async def`` wrappers around ``run_blocking``.
  * Enabled — dispatch the sync function to asyncio's default thread
    pool via ``asyncio.to_thread``. The event loop stays free to serve
    health probes and concurrent requests. Recommended for EKS.

All tool call sites uniformly write ``await run_blocking(func, *args,
**kwargs)`` — the helper internally chooses the dispatch strategy based
on config, so call sites never branch on the flag.

This helper lives in its own module (not inside ``main.py``) so that
dynamically-registered tool modules can import it without creating a
circular dependency back through ``main``.
"""

from __future__ import annotations

import asyncio
import functools
import logging
from typing import Callable, TypeVar

from config import get_config

logger = logging.getLogger(__name__)

T = TypeVar("T")


async def run_blocking(func: Callable[..., T], /, *args, **kwargs) -> T:
    """Run a synchronous callable, optionally on a worker thread.

    Dispatch is governed by ``config.run_blocking_by_asyncio_thread_enabled``
    (env var ``RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED``):

    * **Enabled** — the call is offloaded to asyncio's default thread
      pool via ``asyncio.to_thread``. The event loop remains responsive
      to health probes and concurrent requests while ``func`` runs.
    * **Disabled (default)** — the call runs inline on the event loop,
      preserving the original blocking behaviour. ``func``'s return
      value is returned without ever yielding to the loop.

    The signature is identical in both modes — call sites always write
    ``await run_blocking(func, *args, **kwargs)``.

    Args:
        func: The blocking function to execute.
        *args: Positional arguments forwarded to ``func``.
        **kwargs: Keyword arguments forwarded to ``func``.

    Returns:
        Whatever ``func`` returns.

    Raises:
        Any exception ``func`` raises propagates to the caller. In
        threaded mode the exception is marshalled back from the worker
        thread by ``asyncio.to_thread`` automatically.
    """
    # Read the flag lazily, on every call, so that tests (or admins)
    # that mutate the config singleton between calls see the new value
    # without having to reload this module.
    if get_config().run_blocking_by_asyncio_thread_enabled:
        logger.info("MZA: %s is running by asyncio.to_thread", func.__name__)
        # functools.partial binds kwargs cleanly; asyncio.to_thread accepts
        # *args/**kwargs directly, but going through partial keeps the
        # call site readable and makes the closure easy to log if needed.
        bound = functools.partial(func, *args, **kwargs)
        return await asyncio.to_thread(bound)

    # Legacy path: call inline on the event loop (blocks until done).
    return func(*args, **kwargs)

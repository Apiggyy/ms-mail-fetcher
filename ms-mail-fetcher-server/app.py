import sys

import uvicorn
from app.runtime import create_app, logger, resolve_server_bind


app = create_app()


if __name__ == "__main__":
    host, port, reload_enabled, config_path = resolve_server_bind()

    app.state.server_host = host
    app.state.server_port = port
    logger.info(f"Runtime config loaded: {config_path}")
    logger.info(f"Binding server to {host}:{port}")

    uvicorn.run(
        app,
        host=host,
        port=port,
        reload=reload_enabled and not getattr(sys, "frozen", False),
        access_log=False,
    )

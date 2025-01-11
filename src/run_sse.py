import uvicorn
from starlette.applications import Starlette
from starlette.routing import Route, Mount
from mcp.server.sse import SseServerTransport
from server import create_server
import logging

logger = logging.getLogger(__name__)

class SSEHandler:
    def __init__(self, server, init_options):
        self.server = server
        self.init_options = init_options
        self.sse = SseServerTransport("/messages/")

    async def handle_sse(self, request):
        async with self.sse.connect_sse(
            request.scope, request.receive, request._send
        ) as streams:
            await self.server.run(
                streams[0], streams[1],
                self.init_options
            )

def main():
    server, init_options = create_server()
    sse_handler = SSEHandler(server, init_options)

    routes = [
        Route("/sse", endpoint=sse_handler.handle_sse),
        Mount("/messages/", app=sse_handler.sse.handle_post_message)
    ]

    app = Starlette(routes=routes)
    config = uvicorn.Config(
        app,
        host="0.0.0.0",
        port=8958,
        log_level="debug",
        log_config=None
    )

    server = uvicorn.Server(config)
    try:
        server.run()
    except Exception as e:
        logger.error(f"Failed to start server: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
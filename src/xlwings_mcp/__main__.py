import asyncio
import typer

from .server import run_sse, run_stdio, run_streamable_http

app = typer.Typer(
    help="Excel MCP Server",
    invoke_without_command=True  # 명령 없이 실행 가능
)

@app.callback()
def callback(ctx: typer.Context):
    """
    기본 명령어 없이 실행 시 stdio 모드로 실행
    """
    if ctx.invoked_subcommand is None:
        stdio()

@app.command()
def sse():
    """Start Excel MCP Server in SSE mode"""
    print("Excel MCP Server - SSE mode")
    print("----------------------")
    print("Press Ctrl+C to exit")
    try:
        asyncio.run(run_sse())
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Service stopped.")

@app.command()
def streamable_http():
    """Start Excel MCP Server in streamable HTTP mode"""
    print("Excel MCP Server - Streamable HTTP mode")
    print("---------------------------------------")
    print("Press Ctrl+C to exit")
    try:
        asyncio.run(run_streamable_http())
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Service stopped.")

@app.command()
def stdio():
    """Start Excel MCP Server in stdio mode"""
    print("Excel MCP Server - Stdio mode")
    print("-----------------------------")
    print("Press Ctrl+C to exit")
    try:
        run_stdio()
    except KeyboardInterrupt:
        print("\nShutting down server...")
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("Service stopped.")

if __name__ == "__main__":
    app() 
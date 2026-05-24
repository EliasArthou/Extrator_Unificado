"""
run.py — Entry point para iniciar o ExtratorUnificado Web.

Uso:
    python -m webapp.run                        # porta 8000 (padrão)
    python -m webapp.run --port 8080            # porta customizada
    python -m webapp.run --host 0.0.0.0         # aceita conexões externas
"""

import argparse
import sys
import os

# Garante que a raiz do projeto está no path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import uvicorn


def main():
    parser = argparse.ArgumentParser(description="ExtratorUnificado Web")
    parser.add_argument("--host", default="127.0.0.1", help="Host (padrão: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=8000, help="Porta (padrão: 8000)")
    parser.add_argument("--reload", action="store_true", help="Hot reload para desenvolvimento")
    args = parser.parse_args()

    print(f"\n{'='*60}")
    print(f"  ExtratorUnificado Web")
    print(f"  Acesse: http://{args.host}:{args.port}")
    print(f"  Login padrão: admin@extrator.local / admin123")
    print(f"{'='*60}\n")

    uvicorn.run(
        "webapp.app:app",
        host=args.host,
        port=args.port,
        reload=args.reload,
    )


if __name__ == "__main__":
    main()

import sys

from .web import create_app


if __name__ == "__main__":
    if len(sys.argv) > 1:
        from .cli import main

        raise SystemExit(main())
    create_app().run(host="127.0.0.1", port=5000, debug=False)

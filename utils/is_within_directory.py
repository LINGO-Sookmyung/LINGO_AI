import pathlib

def is_within_directory(child: str, parent: str) -> bool:
    try:
        return pathlib.Path(child).resolve().is_relative_to(pathlib.Path(parent).resolve())
    except AttributeError:
        child_p = pathlib.Path(child).resolve()
        parent_p = pathlib.Path(parent).resolve()
        try:
            child_p.relative_to(parent_p)
            return True
        except Exception:
            return False
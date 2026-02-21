"""COM error formatting."""


def format_com_error(e: Exception) -> str:
    try:
        import pythoncom
        if isinstance(e, pythoncom.com_error):
            hr, msg, exc, arg = e.args
            details = exc[2] if exc else "No details"
            return f"COM Error (0x{hr & 0xFFFFFFFF:08X}): {msg} - {details}"
    except Exception:
        pass
    return str(e)

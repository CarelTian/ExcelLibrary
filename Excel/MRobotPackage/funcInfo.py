import types


def funcInfo(
        name=None,
        argsNote=None,
        note=None,
        methodPath=None
):
    if argsNote is None:
        argsNote = {}

    def decorator(func):
        func.funcPath = methodPath
        func.funcTitle = name
        func.funcArgs = list(argsNote.keys())
        func.funcArgsNote = argsNote
        func.funcNote = note
        return func

    return decorator

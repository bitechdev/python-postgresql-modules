def notice(arguments):
    print("plpy -> notice: ", arguments)


def error(arguments):
    print("plpy -> error: ", arguments)


def warn(arguments):
    print("plpy -> warn: ", arguments)


def execute(arguments):
    print("plpy -> execute: ", arguments)
    return {}

def quote_literal(str):
    print("plpy -> quote_literal: ", str)
    return '"{}"'.format(str)


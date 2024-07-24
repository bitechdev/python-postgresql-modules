def notice(msg, **kwargs):
    print("plpy -> notice: ", msg, kwargs)


def error(msg, **kwargs):
    print("plpy -> error: ", msg, kwargs)


def debug(msg, **kwargs):
    print("plpy -> error: ", msg, kwargs)


def warning(msg, **kwargs):
    print("plpy -> warning: ", msg, kwargs)


def log(msg, **kwargs):
    print("plpy -> log: ", msg, kwargs)


def fatal(msg, **kwargs):
    print("plpy -> fatal: ", msg, kwargs)


def info(msg, **kwargs):
    print("plpy -> info: ", msg, kwargs)


def fatal(msg, **kwargs):
    print("plpy -> fatal: ", msg, kwargs)


def quote_ident(str):
    print("plpy -> quote_ident: ", str)
    return '"{}"'.format(str)


def quote_nullable(str):
    print("plpy -> quote_nullable: ", str)
    return '"{}"'.format(str)


def quote_literal(str):
    print("plpy -> quote_literal: ", str)
    return '"{}"'.format(str)


def execute(arguments):
    print("plpy -> execute: ", arguments)
    row = Row({})
    return row


def prepare(arguments):
    print("plpy -> prepare: ", arguments)
    row = Row({})
    return row


def cursor(arguments):
    print("plpy -> cursor: ", arguments)

    return {}


class SPIError(Exception):
    """Exception for SPIError class"""

    pass


class Row(dict):
    def __init__(self, row_dict: dict):
        """
        :param row_dict: dict containing keys as column names and corresponding values as column values
        """
        self._row_dict = row_dict
        self._set_result_attributes()

    def _set_result_attributes(self):
        """automatically sets each dict key as an attribute and the value as a property
        You shouldn't need to call this directly
        """
        for key, val in self._row_dict.items():
            setattr(self, key, val)

    def __getattr__(self, item):
        return super().getattr(self, item.lower())

    def __setattr__(self, key: str, value):
        if key == "_row_dict":
            super().__setattr__(key, value)
        elif key in self._row_dict.keys():
            self._row_dict[key] = value
            super().__setattr__(key, value)
        else:
            raise RowException(
                "You cannot set {k} to {v} since {k} is not a column in this Row. Row columns: {ks}".format(
                    k=key, v=value, ks=self._row_dict.keys()
                )
            )

    @property
    def row_dict(self):
        """Get the row dictionary at its currents state. Modifying this dictionary directly won't do anything.
        Instead of modifying directly, modify the attribute on the object itself, like so

        >>> from plpy_wrapper import PLPYWrapper
        >>> plpy_wrapper = PLPYWrapper(globals())
        >>> row = plpy_wrapper.execute('select id,name from customer.contact LIMIT 1')[0]
        >>> row.name = 'new name'
        """
        # returning a new dict to protect from direct modification
        return dict(self._row_dict)

    def __repr__(self):
        return self._row_dict.__repr__()

    def __eq__(self, other: "Row"):
        if type(other) is not Row:
            raise NotImplementedError
        elif other.row_dict == self.row_dict:
            return True
        else:
            return False


class ResultSet:
    """wrapper around result of query in plpy https://www.postgresql.org/docs/11/plpython-database.html"""

    def __init__(self, result_set):
        """
        :param result_set: the type expected is the output of plpy.execute, plpy being postgres's native python package
        """
        self.result_set = result_set
        # nrows is the number of rows processes, not number of rows returned by query. Therefore we can't use that variable to iterate. Instead we
        # iterate thru PLyResult object to get all rows and store that in the object
        self._result_set_rows = [row for row in self.result_set]
        self._iterindex = 0

    def __len__(self):
        return len(self.result_set)

    def __str__(self):
        return self.result_set.__str__()

    def __getitem__(self, index: int):
        return Row(self.result_set[index])

    def __iter__(self):
        return self

    def __next__(self) -> Row:
        """this method is here for iteration support"""
        if self._iterindex < len(self._result_set_rows):
            return_val = Row(self._result_set_rows[self._iterindex])
            self._iterindex += 1
            return return_val
        # restart index so we can iterate again next time
        self._iterindex = 0
        raise StopIteration()

    def __repr__(self):
        return "ResultSet=" + str([row for row in self])

    @property
    def n_rows(self) -> int:
        """Returns the number of rows processed by the command"""
        return self.result_set.nrows()

    @property
    def status(self) -> int:
        """The ``SPI_execute()`` return value"""
        return self.result_set.status()

    @property
    def colnames(self):
        """returns a list of column names"""
        return self.result_set.colnames()

    @property
    def coltypes(self):
        """returns list of column type ``OID`` s"""
        return self.result_set.coltypes()

    @property
    def coltypmods(self):
        """returns a list of type-specific type modifiers for the columns"""
        return self.result_set.coltypmods()

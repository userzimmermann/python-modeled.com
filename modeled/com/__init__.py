from six import with_metaclass

__all__ = ['mCOM']

from itertools import chain

from moretools import camelize, decamelize

from win32com.client import DispatchBaseClass
from win32com.client.gencache import EnsureDispatch


class Type(type):
    """Metaclass for modeled :class:`COM` interface.
    """
    def __getitem__(cls, comname):
        com = EnsureDispatch(comname)
        return cls(com)

    def __getattr__(cls, comname):
        return cls[comname]


class COM(with_metaclass(Type, object)):
    """The modeled COM wrapper interface.
    """
    def __init__(self, com):
        self.com = com

    def __getattr__(self, name):
        value = getattr(self.com, camelize(name))
        if isinstance(value, DispatchBaseClass):
            return type(self)(value)
        return value

    def __dir__(self):
        names = set(chain(self.com._prop_map_get_, self.com._prop_map_put_))
        return list(chain(dir(self.com), names, map(decamelize, names)))


mCOM = COM
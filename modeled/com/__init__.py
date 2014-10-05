# python-modeled.com
#
# Copyright (C) 2014 Stefan Zimmermann <zimmermann.code@gmail.com>
#
# python-modeled.com is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# python-modeled.com is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Lesser General Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with python-modeled.com.  If not, see <http://www.gnu.org/licenses/>.

"""modeled.com

.. moduleauthor:: Stefan Zimmermann <zimmermann.code@gmail.com>
"""
from six import with_metaclass

__all__ = ['mCOM']

from itertools import chain

from moretools import camelize, decamelize

from win32com.client import DispatchBaseClass
from win32com.client.gencache import EnsureDispatch

from modeled import mobject


class Type(mobject.type):
    """Metaclass for modeled :class:`COM` interface.
    """
    def __getitem__(cls, comname):
        com = EnsureDispatch(comname)
        return cls(com)

    def __getattr__(cls, comname):
        return cls[comname]


class COM(with_metaclass(Type, mobject)):
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
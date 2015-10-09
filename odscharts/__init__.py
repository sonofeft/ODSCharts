# Python 2 and 3
from __future__ import unicode_literals
from __future__ import absolute_import
from __future__ import print_function
import os

here = os.path.abspath(os.path.dirname(__file__))

exec( open(os.path.join( here,'_version.py' )).read() )  # creates local __version__ variable

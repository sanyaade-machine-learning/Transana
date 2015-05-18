#
# Copyright (C) 2002 The Board of Regents of the University of Wisconsin System 
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of version 2 of the GNU General Public License as
# published by the Free Software Foundation.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software
# Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
#

"""This module implements the WindowControl class as part of the Control
Objects."""

__author__ = 'Nathaniel Case <nacase@wisc.edu>'


class WindowControl(object):
    """This class is responsible for positioning all of the visible windows.
    This is done initially when Presentation Mode or Video Size is activated,
    and when the Media Window is resized if Auto-Arrange is on."""

    def __init__(self):
        """Initialize an WindowControl object."""

# Public methods
    def clear_all_windows(self):
        """Clear all windows and objects."""
    def save_current_positions(self):
        """Save all current window positions."""
    def restore_all_windows(self):
        """Reset all windows to last saved positions."""
    def normal_size(self):
        """Position all windows relative to 100% media size."""
    def one_and_a_half_size(self):
        """Position all windows relative to 150% media size."""
    def double_size(self):
        """Position all windows relative to 200% media size."""
    def setup_presentation(components):
        """Positions all windows for presentation mode.  Components to show
        include video, transcript, or both."""

# Private methods    

# Public properties
    auto_arrange = property(None, None, None,
                        """Should all windows re-position when the Media
                        Window changes?""")

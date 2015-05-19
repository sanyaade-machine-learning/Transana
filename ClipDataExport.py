# Copyright (C) 2003 - 2007 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Clip Data Export function. """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

import wx
# import Python's os module
import os
import Dialogs
import Episode
import Collection
import Clip
import DBInterface
import FilterDialog
import TransanaGlobal
import Misc

class ClipDataExport(Dialogs.GenForm):
    """ This class creates the tab-delimited text file that is the Clip Data Export. """
    def __init__(self, parent, id, episodeNum=0, collectionNum=0):
        # Remember the episode or collection that triggered creation of this report
        self.episodeNum = episodeNum
        self.collectionNum = collectionNum

        # Create a form to get the name of the file to receive the data
        # Define the form title
        title = _("Transana Clip Data Export")
        # Create the form itself
        Dialogs.GenForm.__init__(self, parent, id, title, (550,150), style=wx.DEFAULT_DIALOG_STYLE | wx.RESIZE_BORDER, HelpContext='Clip Data Export')
        # Define the minimum size for this dialog as the initial size
        self.SetSizeHints(550, 150)

        # Export Message Layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.panel, wx.Top, 10)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.height.AsIs()
        # If the filename path is not empty, we need to tell the user.
        prompt = _('Please create a Transana Clip Data File for export.')
        exportText = wx.StaticText(self.panel, -1, prompt)
        exportText.SetConstraints(lay)

        # Export Filename Layout
        lay = wx.LayoutConstraints()
        lay.top.Below(exportText, 10)
        lay.left.SameAs(self.panel, wx.Left, 10)
        lay.width.PercentOf(self.panel, wx.Width, 80)  # 80% width
        lay.height.AsIs()
        self.exportFile = self.new_edit_box(_("Export Filename"), lay, '')
        self.exportFile.SetDropTarget(EditBoxFileDropTarget(self.exportFile))

        # Browse button layout
        lay = wx.LayoutConstraints()
        lay.top.SameAs(self.exportFile, wx.Top)
        lay.left.RightOf(self.exportFile, 10)
        lay.right.SameAs(self.panel, wx.Right, 10)
        lay.bottom.SameAs(self.exportFile, wx.Bottom)
        browse = wx.Button(self.panel, wx.ID_FILE1, _("Browse"), wx.DefaultPosition)
        browse.SetConstraints(lay)
        wx.EVT_BUTTON(self, wx.ID_FILE1, self.OnBrowse)

        self.Layout()
        self.SetAutoLayout(True)
        self.CenterOnScreen()

        self.exportFile.SetFocus()


    def Export(self):
        """ Export the Clip Data to a Tab-delimited file """
        # Determine the appropriate encoding for the export file.  UTF8 does not import into either Excel
        # or SPSS on my computer.  All values other than "latin1" are untested guesses on my part.
        if TransanaGlobal.configData.language == 'ru':
            EXPORT_ENCODING = 'koi8_r'
        elif (TransanaGlobal.configData.language == 'zh'):
            EXPORT_ENCODING = TransanaConstants.chineseEncoding
        elif (TransanaGlobal.configData.language == 'easteurope'):
            EXPORT_ENCODING = 'iso8859_2'
        elif (TransanaGlobal.configData.language == 'el'):
            EXPORT_ENCODING = 'iso8859_7'
        elif (TransanaGlobal.configData.language == 'ja'):
            EXPORT_ENCODING = 'cp932'
        elif (TransanaGlobal.configData.language == 'ko'):
            EXPORT_ENCODING = 'cp949'
        else:
            EXPORT_ENCODING = 'latin1'

        # Initialize values for data structures for this report
        # The Clip List is the list of Clips to be sent to the Filter Dialog
        clipList = []
        # The Clip Lookup allows us to find the Clip Number based on the data from the Clip List
        clipLookup = {}
        # The Keyword List is the list of Keywords to be sent to the Filter Dialog
        keywordList = []

        # If we have an Episode Number, we set up the Episode Clip Data Export
        if self.episodeNum <> 0:
            # First, we get a list of all the Clips for the Episode specified
            tempClipList = DBInterface.list_of_clips_by_episode(self.episodeNum)
            # For all the Clips ...
            for clipRecord in tempClipList:
                # ... add the Clip to the Clip List for filtering ...
                clipList.append((clipRecord['ClipID'], clipRecord['CollectNum'], True))
                # ... retain a pointer to the Clip Number keyed to the Clip ID and Collection Number ...
                clipLookup[(clipRecord['ClipID'], clipRecord['CollectNum'])] = clipRecord['ClipNum']
                # ... now get all the keywords for this Clip ...
                clipKeywordList = DBInterface.list_of_keywords(Clip = clipRecord['ClipNum'])
                # ... and iterate through the list of clip keywords.
                for clipKeyword in clipKeywordList:
                    # If the keyword isn't already in the Keyword List ...
                    if (clipKeyword[0], clipKeyword[1], True) not in keywordList:
                        # ... add the keyword to the keyword list for filtering.
                        keywordList.append((clipKeyword[0], clipKeyword[1], True))

        # If we don't have an Episode number, but DO have a Collection Number, we set up the Clips for the Collection specified
        elif self.collectionNum <> 0:
            # First, load the specified collection.  We need its data.
            tempCollection = Collection.Collection(self.collectionNum)
            # Get a list of all teh Clips for the Collection specified.
            tempClipList = DBInterface.list_of_clips_by_collection(tempCollection.id, tempCollection.parent)
            # For all the Clips ...
            for (clipNo, clipName, collNo) in tempClipList:
                # ... add the Clip to the Clip List for filtering ...
                clipList.append((clipName, collNo, True))
                # ... retain a pointer to the Clip Number keyed to the Clip ID and Collection Number ...
                clipLookup[(clipName, collNo)] = clipNo
                # ... now get all the keywords for this Clip ...
                clipKeywordList = DBInterface.list_of_keywords(Clip = clipNo)
                # ... and iterate through the list of clip keywords.
                for clipKeyword in clipKeywordList:
                    # If the keyword isn't already in the Keyword List ...
                    if (clipKeyword[0], clipKeyword[1], True) not in keywordList:
                        # ... add the keyword to the keyword list for filtering.
                        keywordList.append((clipKeyword[0], clipKeyword[1], True))

        # Put the Clip List in alphabetical order in preparation for Filtering..
        clipList.sort()
        # Put the Keyword List in alphabetical order in preparation for Filtering.
        keywordList.sort()

        # Prepare the Filter Dialog.
        # Set the title for the Filter Dialog
        title = _("Clip Data Export Filter Dialog")
        # If we have an Episode-based report ...
        if self.episodeNum != 0:
            # ... reportType 3 indicates Episode Clip Data Export to the Filter Dialog
            reportType = 3
            # ... the reportScope is the Episode Number.
            reportScope = self.episodeNum
        # If we have a Collection-based report ...
        else:
            # ... reportType 4 indicates Collection Clip Data Export to the Filter Dialog
            reportType = 4
            # ... the reportScope is the Collection Number.
            reportScope = self.collectionNum
        
        # Create a Filter Dialog, passing all the necessary parameters.
        dlgFilter = FilterDialog.FilterDialog(None, -1, title, reportType=reportType, reportScope=reportScope,
                                              clipFilter=True, keywordFilter=True, keywordSort=False,
                                              options=False)
        # Populate the Clip and Keyword Data Structures
        dlgFilter.SetClips(clipList)
        dlgFilter.SetKeywords(keywordList)

        # If the user clicks OK ...
        if dlgFilter.ShowModal() == wx.ID_OK:
            TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
            # Get the revised Clip and Keyword data from the Filter Dialog
            clipList = dlgFilter.GetClips()
            keywordList = dlgFilter.GetKeywords()

            # Get the user-specified File Name
            fs = self.exportFile.GetValue()
            # Ensure that the file name has the proper extension.
            if fs[-4:].lower() != '.txt':
                fs = fs + '.txt'
            # On the Mac, if no path is specified, the data is exported to a file INSIDE the application bundle, 
            # where no one will be able to find it.  Let's put it in the user's HOME directory instead.
            # I'm okay with not handling this on Windows, where it will be placed in the Program's folder
            # but it CAN be found.  (There's no easy way on Windows to determine the location of "My Documents"
            # especially if the user has moved it.)
            if "__WXMAC__" in wx.PlatformInfo:
                # if the specified file has no path specification ...
                if fs.find(os.sep) == -1:
                    # ... then prepend the HOME folder
                    fs = os.getenv("HOME") + os.sep + fs
            # Open the output file for writing.
            f = file(fs, 'w')

            # Write the Header line.  We're creating a tab-delimited file, so we'll use tabs to separate the items.
            f.write(_('Collection Name\tClip Name\tMedia File\tClip Start\tClip End\tClip Length (seconds)'))
            # Add keywords to the Header.  Iterate through the Keyword List.
            for keyword in keywordList:
                # See if the user has left the keyword "checked" in the filter dialog.
                if keyword[2]:
                    # Encode and write all "checked" keywords to the Header.
                    kwg = keyword[0].encode(EXPORT_ENCODING)
                    kw = keyword[1].encode(EXPORT_ENCODING)
                    f.write('\t%s : %s' % (kwg, kw))
            # Add a line break to signal the end of the Header line. 
            f.write('\n')

            # Now iterate through the Clip List
            for clipRec in clipList:
                # See if the user has left the clip "checked" in the filter dialog.
                if clipRec[2]:
                    # Load the Clip data.  The ClipLookup dictionary allows this easily.
                    clip = Clip.Clip(clipLookup[clipRec[0], clipRec[1]])
                    # Encode string values using the Export Encoding
                    collectionID = clip.collection_id.encode(EXPORT_ENCODING)
                    clipID = clip.id.encode(EXPORT_ENCODING)
                    clipMediaFilename = clip.media_filename.encode(EXPORT_ENCODING)
                    # Write the Clip's data values to the output file.  We're creating a tab-delimited file,
                    # so we'll use tabs to separate the items.
                    f.write('%s\t%s\t%s\t%s\t%s\t%10.4f' % (collectionID, clipID, clipMediaFilename,
                                                        Misc.time_in_ms_to_str(clip.clip_start), Misc.time_in_ms_to_str(clip.clip_stop),
                                                        (clip.clip_stop - clip.clip_start) / 1000.0))

                    # Now we iterate through the keyword list ...
                    for keyword in keywordList:
                        # ... looking only at those keywords the user left "checked" in the filter dialog ...
                        if keyword[2]:
                            # ... and check to see if the Clip HAS the keyword.
                            if clip.has_keyword(keyword[0], keyword[1]):
                                # If so, we write a "1", indicating True.
                                f.write('\t1')
                            else:
                                # If not, we write a "0", indicating False.
                                f.write('\t0')
                    # Add a line break to signal the end of the Clip record
                    f.write('\n')

            # Flush the output file's buffer (probably unnecessary)
            f.flush()
            # Close the output file.
            f.close()
            TransanaGlobal.menuWindow.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
            # If so, create a prompt to inform the user and ask to overwrite the file.
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Clip Data has been exported to file "%s".'), 'utf8')
            else:
                prompt = _('Clip Data has been exported to file "%s".')
            # Create the dialog to inform the user
            dlg2 = Dialogs.InfoDialog(self, prompt % fs)
            # Show the Info dialog.
            dlg2.ShowModal()
            # Destroy the Information dialog
            dlg2.Destroy()                    

        # Destroy the Filter Dialog.  We're done with it.
        dlgFilter.Destroy()
        
    def OnBrowse(self, evt):
        """Invoked when the user activates the Browse button."""
        fs = wx.FileSelector(_("Select a text file for export"),
                        TransanaGlobal.programDir,
                        "",
                        "", 
                        _("Text Files (*.txt)|*.txt|All files (*.*)|*.*"), 
                        wx.SAVE)
        # If user didn't cancel ..
        if fs != "":
            self.exportFile.SetValue(fs)


# This simple derrived class let's the user drop files onto an edit box
class EditBoxFileDropTarget(wx.FileDropTarget):
    def __init__(self, editbox):
        wx.FileDropTarget.__init__(self)
        self.editbox = editbox
    def OnDropFiles(self, x, y, files):
        """Called when a file is dragged onto the edit box."""
        self.editbox.SetValue(files[0])

# Copyright (C) 2003 - 2015 The Board of Regents of the University of Wisconsin System 
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

"""This module implements the Control Object class for Transana,
which is responsible for managing communication between the
four main windows.  Each object (Menu, Visualization, Video, Transcript,
and Data) should communicate only with the Control Object, not with
each other.
"""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Rajas Sambhare'

DEBUG = False
if DEBUG:
    print "ControlObjectClass DEBUG is ON!"

# Import wxPython
import wx

# import Transana's Constants
import TransanaConstants
# Import the Menu Constants
import MenuSetup
# Import Transana's Global Values
import TransanaGlobal
# import the Transana Library Object definition
import Library
# import the Transana Episode Object definition
import Episode
# import the Transana Transcript Object definition
import Transcript
# import the Transana Document Object definition
import Document
# import the Transana Collection Object definition
import Collection
# import the Transana Clip Object definition
import Clip
# import teh Transana Quote Object definition
import Quote
# import the Transana Miscellaneous Routines
import Misc
# import Transana Database Interface
import DBInterface
# import Transana's Dialogs
import Dialogs
# import Transana's DragAndDrop Objects for Quick Clip creation
import DragAndDropObjects
# import Transana File Management System
import FileManagement
# import Play All Clips
import PlayAllClips
# import the Episode Transcript Change Propagation tool
import PropagateChanges
# import Transana's Snapshot object
import Snapshot
# import the Snapshot Window
import SnapshotWindow
# import Transana's Exceptions
import TransanaExceptions
# Import Transana's Transcript User Interface for creating supplemental Transcript Windows
if TransanaConstants.USESRTC:
    import TranscriptionUI_RTC as TranscriptionUI
else:
    import TranscriptionUI
# import Python's os module
import os
# import Python's sys module
import sys
# import Python's string module
import string
# Import Python's fast cPickle module
import cPickle
# import Python's pickle module
import pickle


class ControlObject(object):
    """ The ControlObject operationalizes all inter-window and inter-object communication and control.
        All objects should speak only to the ControlObject, not to each other directly.  The purpose of
        this is to allow greater modularity of code, so that modules can be swapped in and out in with
        changes affecting only this object if the APIs change.  """
    def __init__(self):
        """ Initialize the ControlObject """
        # Define Objects that need controlling (initializing to None)
        self.MenuWindow = None
        self.VideoWindow = None
        self.TranscriptWindow = None
        self.shuttingDown = False       # We need to signal when we want to shut down to prevent problems
                                        # with the Visualization Window's IDLE event trying to call the
                                        # VideoWindow after it's been destroyed.
        # We need to know what transcript is "Active" (most recently selected) at any given point.  -1 signals none.
        self.activeTranscript = -1
        self.VisualizationWindow = None
        self.DataWindow = None
        # Keep track of all Snapshot Windows that are opened
        self.SnapshotWindows = []
        self.PlayAllClipsWindow = None
        self.NotesBrowserWindow = None
        self.ChatWindow = None
        # Keep track of all Report, Map, and Graph Windows that are opened
        self.ReportWindows = {}

        # Initialize variables
        self.VideoFilename = ''         # Video File Name
        self.VideoStartPoint = 0        # Starting Point for video playback in Milliseconds
        self.VideoEndPoint = 0          # Ending Point for video playback in Milliseconds
        self.WindowPositions = []       # Initial Screen Positions for all Windows, used for Presentation Mode
        self.TranscriptNum = {}         # Transcript Num is a dictionary with Transcript Number as key and (Tab, Pane) as data
        self.currentObj = None          # Currently loaded Object (Episode or Clip)
        self.reportNumber = 0           # Report Number, for tracking reports in the Window Menu
        # Have the Export Directory default to the Video Root, but then remember its changed value for the session
        self.defaultExportDir = TransanaGlobal.configData.videoPath
        self.playInLoop = False         # Should we loop playback?
        self.LoopPresMode = None        # What presentation mode are we ignoring while Looping?
        self.shutdownPlayAllClips = False  # Flag to signal the need to reformat the screen following Play All Clips

    def Register(self, Menu='', Video='', Transcript='', Data='', Visualization='', PlayAllClips='', NotesBrowser='', Chat=''):
        """ The ControlObject can extert control only over those objects it knows about.  This method
            provides a way to let the ControlObject know about other objects.  This infrastructure allows
            for objects to be swapped in and out.  For example, if you need a different video window
            that supports a format not available on the current one, you can hide the current one, show
            a new one, and register that new one with the ControlObject.  Once this is done, the new
            player will handle all tasks for the program.  """
        # This function expects parameters passed by name and "registers" the components that
        # need to be available to the ControlObject to be controlled.  To remove an
        # object registration, pass in "None"
        if Menu != '':
            self.MenuWindow = Menu                       # Define the Menu Window Object
        if Video != '':
            self.VideoWindow = Video                     # Define the Video Window Object
        if Transcript != '':
            # Define the Transcript Window Object
            self.TranscriptWindow = Transcript
##            # Add the Transcript Number to the list of Transcript Numbers
##            self.TranscriptNum.append(0)
##            # Set the new Transcript to be the Active Transcript
##            self.activeTranscript = len(self.TranscriptWindow) - 1
        if Data != '':
            self.DataWindow = Data                       # Define the Data Window Object
        if Visualization != '':
            self.VisualizationWindow = Visualization     # Define the Visualization Window Object
        if PlayAllClips != '':
            self.PlayAllClipsWindow = PlayAllClips       # Define the Play All Clips Window Object
        if NotesBrowser != '':
            self.NotesBrowserWindow = NotesBrowser             # Define the Notes Browser Window Object
        if Chat != '':
            self.ChatWindow = Chat                       # Define the Chat Window Object

    def CloseAll(self):
        """ This method closes all application windows and cleans up objects when the user
            quits Transana. """
        # Closing the MenuWindow will automatically close the Transcript, Data, and Visualization
        # Windows in the current setup of Transana, as these windows are all defined as child dialogs
        # of the MenuWindow.
        self.MenuWindow.Close()
        # VideoWindow needs to be closed explicitly.
        self.VideoWindow.close()

    def CloseCurrentTranscript(self, event):
        """ Close the current Transcript Window """
        # Have the Transcript Window close the current Pane or Panel
        self.TranscriptWindow.CloseCurrent(event)

    def CloseAllImages(self):
        """ Close all Snapshot Windows """
        # For each Shapshot Window (from the end of the list to the start) ...
        while len(self.SnapshotWindows) > 0:
            # ... close it, thus releasing any records that might be locked there.
            self.SnapshotWindows[len(self.SnapshotWindows) - 1].Close()

    def CloseAllReports(self):
        """ Close all Report Windows """
        # For each Report Window (from the end of the list to the start) ...
        while len(self.ReportWindows) > 0:
            # ... close it, thus releasing any records that might be locked there.
            self.ReportWindows[self.ReportWindows.keys()[len(self.ReportWindows) - 1]].Close()

    def IconizeAll(self, iconize):
        """ Have all windows minimize and restore together """
        self.MenuWindow.Iconize(iconize)
        self.VisualizationWindow.Iconize(iconize)
        self.VideoWindow.Iconize(iconize)
        # The TranscriptWindow sometimes MUST be called here, while other times it isn't needed.
        self.TranscriptWindow.Iconize(iconize)
        self.DataWindow.Iconize(iconize)
        for win in self.SnapshotWindows:
            win.Iconize(iconize)
        if self.NotesBrowserWindow != None:
            self.NotesBrowserWindow.Iconize(iconize)
        # The File Management Window also does not need to be processed here.
        # For each Report Window ...
        for win in self.ReportWindows.keys():
            # ... minimize/restore the Report
            self.ReportWindows[win].Iconize(iconize)

    def LoadDocument(self, library_name, document_name, document_number):
        """ When a Document is identified to trigger systemic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """
        # Initialize a variable indicating if we found the requested document
        documentFound = False
        # First, see if the selected Document is already loaded!  Iterate through the TranscriptWindow's Notebook Tabs ...
        for y in range(self.TranscriptWindow.nb.GetPageCount()):
            for pane in self.TranscriptWindow.nb.GetPage(y).GetChildren():
                # ... and get a pointer to the tab's active Splitter panel's editor's data object
                dataObj = pane.editor.TranscriptObj
                # If the data object is not None, then something IS loaded
                if dataObj is not None:
                    # If the data object is a Document (not a Transcript) and the Document has the same NAME ...
                    if isinstance(dataObj, Document.Document) and (document_name == dataObj.id):
                        # ... load the data object's Library
                        library = Library.Library(dataObj.library_num)
                        # Also load the Database copy of this Document, so we can check that it's up to date
                        # and hasn't been updated by another user.
                        dbDataObj = Document.Document(document_number)
                        # If that library name matches the one we're opening ...
                        if (library_name == library.id):
                            # ... then the requested Document is already open.  Select its Notebook Page ...
                            self.TranscriptWindow.nb.SetSelection(y)
                            # ... and select the correct Splitter Pane as the "Active" pane.
                            self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(pane.panelNum)
                            # Note that the document was found
                            documentFound = True
                            # We can stop looking now
                            break
            # If the document is found ...
            if documentFound:
                # ... we can interrupt this loop too!
                break

        # If the requested document was not found, or if it is not CURRENT (if another user has edited it!) ...
        if not documentFound or ((dataObj != None) and (dataObj.lastsavetime != dbDataObj.lastsavetime)):

            # If the document was not found but the current page is not empty ...
            if not documentFound and (dataObj != None):
                # ... create a new Notebook Page for the Document
                self.TranscriptWindow.AddNotebookPage(document_name)
                # Select the new page as the current page
                self.TranscriptWindow.nb.SetSelection(self.TranscriptWindow.nb.GetPageCount() - 1)
            # If the current Document Window's Notebook Page IS empty ...
            else:
                self.TranscriptWindow.nb.SetPageText(self.TranscriptWindow.nb.GetSelection(), document_name)
                
            # Load the Document
            tmpDocument = Document.Document(document_number)
            # Load the Document into the Editor Interface (Transcripts and Documents act the same here!)
            self.TranscriptWindow.LoadTranscript(tmpDocument)
            # Set the new Current Object
            self.currentObj = tmpDocument

            # If we don't yet know the Document Length ...
            if (self.currentObj.document_length == 0) and (self.TranscriptWindow.dlg.editor.GetLength() > 0):
                # ... let's grab it here and save it.  Otherwise, Visualizations don't work, etc.
                # ... lock the document ...
                self.currentObj.lock_record()
                # ... add the length ...
                self.currentObj.document_length = self.TranscriptWindow.dlg.editor.GetLength()
                # ... save the record ...
                self.currentObj.db_save()
                # ... and unlock the record
                self.currentObj.unlock_record()

            # Update the Transana Interface for this object
            self.UpdateCurrentObject(tmpDocument)
            # And tell the Visualization Window to draw itself.
            self.VisualizationWindow.Refresh()

            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

    def LoadTranscript(self, library, episode, transcript):
        """ When a Transcript is identified to trigger systemic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """
        # First, let's see if there's already a video loaded in the system.  Iterate through all Notebook Pages.
        self.BringTranscriptToFront()
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow.TranscriptModified():
            if TransanaConstants.partialTranscriptEdit:
                self.SaveTranscript(1, cleardoc=1, continueEditing=False)
            else:
                self.SaveTranscript(1, cleardoc=1)
        # If the current Editor is a Document (not None, not a Transcript) ...
        if isinstance(self.TranscriptWindow.GetCurrentObject(), Document.Document) or \
           isinstance(self.TranscriptWindow.GetCurrentObject(), Quote.Quote):
            # ... create a new Notebook Page for the Document
            self.TranscriptWindow.AddNotebookPage(transcript)
            # Select the new page as the current page
            self.TranscriptWindow.nb.SetSelection(self.TranscriptWindow.nb.GetPageCount() - 1)
        # If the current Editor is a Transcript (not None, not a Document) ...
        elif isinstance(self.TranscriptWindow.GetCurrentObject(), Transcript.Transcript):
            # ... then we need to Clear all Windows of media information
            self.ClearAllWindows(clearAllPanes=True)
            if self.currentObj != None:
                # ... create a new Notebook Page for the Document
                self.TranscriptWindow.AddNotebookPage(transcript)
                # Select the new page as the current page
                self.TranscriptWindow.nb.SetSelection(self.TranscriptWindow.nb.GetPageCount() - 1)
        # Because transcript names can be identical for different episodes in different Library, all parameters are mandatory.
        # They are:
        #   Library     -  the Library associated with the desired Transcript
        #   episode     -  the Episode associated with the desired Transcript
        #   transcript  -  the Transcript to be displayed in the Transcript Window
        libraryObj = Library.Library(library)                                    # Load the Library which owns the Episode which owns the Transcript
        episodeObj = Episode.Episode(series=libraryObj.id, episode=episode)   # Load the Episode in the Library that owns the Transcript
        # Set the current object to the loaded Episode
        self.currentObj = episodeObj
        transcriptObj = Transcript.Transcript(transcript, ep=episodeObj.number)

        # Load the Transcript in the Episode in the Library
        # reset the video start and end points
        self.VideoStartPoint = 0                                     # Set the Video Start Point to the beginning of the video
        self.VideoEndPoint = 0                                       # Set the Video End Point to 0, indicating that the video should not end prematurely
        
        # Remove any tabs in the Data Window beyond the Database Tab
        self.DataWindow.DeleteTabs()

        if self.LoadVideo(self.currentObj):    # Load the video identified in the Episode
            # Delineate the appropriate start and end points for Video Control.  (Required to prevent Waveform Visualization problems)
            self.SetVideoSelection(0, 0)

            # Force the Visualization to load here.  This ensures that the Episode visualization is shown
            # rather than the Clip visualization when Locating a Clip
            self.VisualizationWindow.OnIdle(None)
            
            # If we have only one video file ...
            if len(self.currentObj.additional_media_files) == 0:
                # Identify the loaded media file
                prompt = unicode(_('Video Media File: "%s"'), 'utf8')
                # Place the file name in the video window's Title bar
                self.VideoWindow.SetTitle(prompt % episodeObj.media_filename)
            # If there are multiple videos ...
            else:
                # Just label the video window generically.  There's not room for file names.
                self.VideoWindow.SetTitle(_("Media"))
            # Open Transcript in Transcript Window
            self.TranscriptWindow.LoadTranscript(transcriptObj) #flies off to transcriptionui.py

            self.currentObj = episodeObj

            # Update the Transana Interface for this object
            self.UpdateCurrentObject(transcriptObj)

            # Add the Transcript Number to the list that tracks the numbers of the open transcripts
            self.TranscriptNum[transcriptObj.number] = (self.TranscriptWindow.nb.GetSelection(), self.TranscriptWindow.nb.GetPage(self.TranscriptWindow.nb.GetSelection()).activePanel)
            
##            # Add the Episode Clips Tab to the DataWindow
##            self.DataWindow.AddItemsTab(libraryObj=libraryObj, dataObj=episodeObj)
##
##            # Add the Selected Episode Clips Tab, initially set to the beginning of the video file
##            # TODO:  When the Transcript Window updates the selected text, we need to update this tab in the Data Window!
##            self.DataWindow.AddSelectedItemsTab(libraryObj=libraryObj, dataObj=episodeObj, timeCode=0)
##
##            # Add the Keyword Tab to the DataWindow
##            self.DataWindow.AddKeywordsTab(seriesObj=libraryObj, episodeObj=episodeObj)
            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

            if TransanaConstants.USESRTC:
                # After two seconds, call the EditorPaint method of the Transcript Dialog (in the TranscriptionUI_RTC file)
                # This causes improperly placed line numers to "correct" themselves!
                wx.CallLater(2000, self.TranscriptWindow.dlg.EditorPaint, None)

             # Set focus to the new Transcript's Editor (so that CommonKeys work on the Mac)
            self.TranscriptWindow.dlg.editor.SetFocus()

        # If the video won't load ...
        else:
            # Clear the interface!
            self.ClearAllWindows(clearAllPanes=True)
            # We only want to load the File Manager in the Single User version.  It's not the appropriate action
            # for the multi-user version!
            if TransanaConstants.singleUserVersion:
                # Open the File Management Window just as if the Menu Item was selected, which
                # doesn't cause menu problems on OS X
                self.MenuWindow.OnFileManagement(None)

    def GetCurrentItemType(self):
        """ Report whether the currently-selected Item is None (nothing loaded), a Document, or a Transcript """
        # If no object is loaded in the Transcript Window ...
        if self.TranscriptWindow.dlg.editor.TranscriptObj is None:
            # ... we have nothing
            return None
        # If the currently-selected object is a Document ...
        elif isinstance(self.TranscriptWindow.dlg.editor.TranscriptObj, Document.Document):
            # ... we have a Document
            return 'Document'
        # If the currently-selected object is a Transcript ...
        elif isinstance(self.TranscriptWindow.dlg.editor.TranscriptObj, Transcript.Transcript):
            # ... we have a Transcript
            return 'Transcript'
        # If the currently-selected object is a Quote...
        elif isinstance(self.TranscriptWindow.dlg.editor.TranscriptObj, Quote.Quote):
            # ... we have a Quote
            return 'Quote'

    def LoadQuote(self, quote_number):
        """ When a Quote is identified to trigger systemic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """
        # Initialize a variable indicating if we found the requested Quote
        quoteFound = False
        # First, see if the selected Quote is already loaded!  Iterate through the TranscriptWindow's Notebook Tabs ...
        for y in range(self.TranscriptWindow.nb.GetPageCount()):
            for pane in self.TranscriptWindow.nb.GetPage(y).GetChildren():
                # ... and get a pointer to the tab's active Splitter panel's editor's data object
                dataObj = pane.editor.TranscriptObj
                # If the data object is not None, then something IS loaded
                if dataObj is not None:
                    # If the data object is a Quote not a Transcript) and the Quote has the same NUMBER ...
                    if isinstance(dataObj, Quote.Quote) and (quote_number == dataObj.number):
                        # ... then the requested Quote is already open.  Select its Notebook Page ...
                        self.TranscriptWindow.nb.SetSelection(y)
                        # ... and select the correct Splitter Pane as the "Active" pane.
                        self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(pane.panelNum)
                        # Note that the quote was found
                        quoteFound = True
                        # We can stop looking now
                        break                

        # If the requested document was not found ...
        if not quoteFound:
            # Load the Quote
            tmpQuote = Quote.Quote(quote_number)
            # If the current Transcript Window's Notebook Page is NOT empty ...
            if (self.TranscriptWindow.dlg.editor.TranscriptObj != None):
                # ... create a new Notebook Page for the Quote
                self.TranscriptWindow.AddNotebookPage(tmpQuote.id)
                # Select the new page as the current page
                self.TranscriptWindow.nb.SetSelection(self.TranscriptWindow.nb.GetPageCount() - 1)
            # If the current Document Window's Notebook Page IS empty ...
            else:
                self.TranscriptWindow.nb.SetPageText(self.TranscriptWindow.nb.GetSelection(), tmpQuote.id)
                
            # Load the Quote into the Editor Interface (Transcripts, Documents, and Quotes act the same here!)
            self.TranscriptWindow.LoadTranscript(tmpQuote)

##            # Remove any tabs in the Data Window beyond the Database Tab
##            self.DataWindow.DeleteTabs()
##
##            # Load the data object's Library
##            tmpCollection = Collection.Collection(tmpQuote.collection_num)
##            # Add the Keyword Tab to the DataWindow
##            self.DataWindow.AddKeywordsTab(collectionObj = tmpCollection, quoteObj = tmpQuote)
##
##            # Set the new Current Object
##            self.currentObj = tmpQuote
##            # Set the Visualization Window's Visualization Object
##            self.VisualizationWindow.SetVisualizationObject(tmpQuote)
##
            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

            # Update the Transana Interface for this object
            self.UpdateCurrentObject(tmpQuote)

        # Get the current selection(s) from the Database Tree
        selItems = self.DataWindow.DBTab.tree.GetSelections()
        # If there are one or more items selected ...
        if len(selItems) >= 1:
            # ... get the item data from the first selection
            selData = self.DataWindow.DBTab.tree.GetPyData(selItems[0])
        # If NO items are selected ...
        else:
            # ... then there's no item data to get
            selData = None

        # If no items are selected or the item selected is NOT a Search Collection or Search Clip ...
        if (selData == None) or not (selData.nodetype in ['SearchCollectionNode', 'SearchClipNode']):
            # Let's make sure this clip is displayed in the Database Tree
            nodeList = (_('Collections'),) + self.currentObj.GetNodeData()
            if isinstance(self.currentObj, Quote.Quote):
                # Now point the DBTree (the notebook's parent window's DBTab's tree) to the loaded Quote
                self.DataWindow.DBTab.tree.select_Node(nodeList, 'QuoteNode')

    def LoadClipByNumber(self, clipNum):
        """ When a Clip is identified to trigger systematic loading of all related information,
            this method should be called so that all Transana Objects are set appropriately. """

        # Load the Clip based on the ClipNumber.  (Let's get NotFound exceptions out of the way early!)
        clipObj = Clip.Clip(clipNum)

        # First, let's see if there's already a video loaded in the system.  Iterate through all Notebook Pages.
        self.BringTranscriptToFront()
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow.TranscriptModified():
            if TransanaConstants.partialTranscriptEdit:
                self.SaveTranscript(1, cleardoc=1, continueEditing=False)
            else:
                self.SaveTranscript(1, cleardoc=1)
        # If the current Editor is a Document (not None, not a Transcript) ...
        if isinstance(self.TranscriptWindow.GetCurrentObject(), Document.Document) or \
           isinstance(self.TranscriptWindow.GetCurrentObject(), Quote.Quote):
            # ... create a new Notebook Page for the Document
            self.TranscriptWindow.AddNotebookPage(_("No Document Loaded"))
            # Select the new page as the current page
            self.TranscriptWindow.nb.SetSelection(self.TranscriptWindow.nb.GetPageCount() - 1)
        # If the current Editor is a Transcript (not None, not a Document) ...
        elif isinstance(self.TranscriptWindow.GetCurrentObject(), Transcript.Transcript):
            # ... then we need to Clear all Windows of media information
            self.ClearAllWindows(clearAllPanes=True)
            if self.currentObj != None:
                # ... create a new Notebook Page for the Document
                self.TranscriptWindow.AddNotebookPage(clipObj.id)
                # Select the new page as the current page
                self.TranscriptWindow.nb.SetSelection(self.TranscriptWindow.nb.GetPageCount() - 1)
        
        # Set the current object to the loaded Clip
        self.currentObj = clipObj
        # Load the Collection that contains the loaded Clip
        collectionObj = Collection.Collection(clipObj.collection_num)
        # set the video start and end points to the start and stop points defined in the clip
        self.VideoStartPoint = clipObj.clip_start                     # Set the Video Start Point to the Clip beginning
        self.VideoEndPoint = clipObj.clip_stop                        # Set the Video End Point to the Clip end

        # Load the video identified in the Clip
        if self.LoadVideo(self.currentObj):
            # If we have only one video file ...
            if len(self.currentObj.additional_media_files) == 0:
                # Identify the loaded media file
                prompt = unicode(_('Video Media File: "%s"'), 'utf8')
                # Place the file name in the video window's Title bar
                self.VideoWindow.SetTitle(prompt % clipObj.media_filename)
            # If there are multiple videos ...
            else:
                # Just label the video window generically.  There's not room for file names.
                self.VideoWindow.SetTitle(_("Media"))

            # Open the first Clip Transcript in Transcript Window (activeTranscript is ALWAYS 0 here!)
            self.TranscriptWindow.LoadTranscript(clipObj.transcripts[0])
            # Update the Transana Interface for this object
            self.UpdateCurrentObject(clipObj.transcripts[0])
            
            # If we allow multiple transcripts ...
            if TransanaConstants.proVersion:
                # Open the remaining clip transcripts in additional transcript windows.
                for tr in clipObj.transcripts[1:]:
                    self.OpenAdditionalTranscript(tr.number, isEpisodeTranscript=False)

            # Delineate the appropriate start and end points for Video Control
            self.SetVideoSelection(self.VideoStartPoint, self.VideoEndPoint)

            # For reasons I have not been able to track down, when you load a Collection Report, then
            # use Hyperlink to open a Quote, then use Hyperlink to open a Clip, currentObj gets wiped out!
            # This replaces it (again).  I've run the debugger, and the currentObj disappears in a place
            # that makes *NO* sense at all.  I'm baffled.
            if (self.currentObj == None) and (clipObj != None):
                self.currentObj = clipObj

##            # Remove any tabs in the Data Window beyond the Database Tab.  (This was moved down to late in the
##            # process due to problems on the Mac documented in the DataWindow object.)
##            self.DataWindow.DeleteTabs()
##            # Add the Keyword Tab to the DataWindow
##            self.DataWindow.AddKeywordsTab(collectionObj=collectionObj, clipObj=clipObj)
##
##            # Get the current selection(s) from the Database Tree
##            selItems = self.DataWindow.DBTab.tree.GetSelections()
##            # If there are one or more items selected ...
##            if len(selItems) >= 1:
##                # ... get the item data from the first selection
##                selData = self.DataWindow.DBTab.tree.GetPyData(selItems[0])
##            # If NO items are selected ...
##            else:
##                # ... then there's no item data to get
##                selData = None
##
##            # If no items are selected or the item selected is NOT a Search Collection or Search Clip ...
##            if (selData == None) or not (selData.nodetype in ['SearchCollectionNode', 'SearchClipNode']):
##                # Let's make sure this clip is displayed in the Database Tree
##                nodeList = (_('Collections'),) + self.currentObj.GetNodeData()
##                if isinstance(self.currentObj, Clip.Clip):
##                    # Now point the DBTree (the notebook's parent window's DBTab's tree) to the loaded Clip
##                    self.DataWindow.DBTab.tree.select_Node(nodeList, 'ClipNode')

            # Enable the transcript menu item options
            self.MenuWindow.SetTranscriptOptions(True)

            return True
        else:
            # Remove any tabs in the Data Window beyond the Database Tab
            self.DataWindow.DeleteTabs()

            # We only want to load the File Manager in the Single User version.  It's not the appropriate action
            # for the multi-user version!
            if TransanaConstants.singleUserVersion:
                # Create a File Management Window
                fileManager = FileManagement.FileManagement(self.MenuWindow, -1, _("Transana File Management"))
                # Set up, display, and process the File Management Window
                fileManager.Setup(showModal=True)
                # Destroy the File Manager window
                fileManager.Destroy()

            return False

    def LocateQuoteInDocument(self, quoteNum):
        """ Locate the specificed Quote in its source Document, if possible """
        # Load the Quote
        quote = Quote.Quote(quoteNum)
        try:
            # If the Source Document is known ...
            if quote.source_document_num > 0:
                # Load the Document
                document = Document.Document(quote.source_document_num)
                # Load the Library
                library = Library.Library(document.library_num)
                # Load the document into the TranscriptWindow via the ControlObject
                self.LoadDocument(library.id, document.id, quote.source_document_num)
                # Highlight the Quote's text
                self.HighlightQuoteInCurrentDocument(quote)
            else:
                msg = _('The Document this Quote was created from cannot be loaded.\nMost likely, the source Document has been deleted.')
                dlg = Dialogs.ErrorDialog(None, msg)
                result = dlg.ShowModal()
                dlg.Destroy()
        except TransanaExceptions.RecordNotFoundError, e:
            msg = _('The Document this Quote was created from cannot be loaded.\nMost likely, the source Document has been deleted.')
            dlg = Dialogs.ErrorDialog(None, msg)
            result = dlg.ShowModal()
            dlg.Destroy()

    def LocateClipInEpisode(self, clipNum):
        """ Locate the specified Clip in its source Episode, if possible """
        # Load the Clip.  We DO need the Clip Transcript(s) here
        clip = Clip.Clip(clipNum)
        # We need to track what this clip's source transcripts are.
        # Initialize a list to store them.
        sourceTranscripts = []
        # For each clip transcripts ...
        for tr in clip.transcripts:
            # ... add that transcript's source transcript to the list
            sourceTranscripts.append(tr.source_transcript)
        # Start exception handling to catch failures due to orphaned clips
        try:
            # Load the Episode
            episode = Episode.Episode(clip.episode_num)
            # If all source transcripts are KNOWN ...
            if not 0 in sourceTranscripts:
                # ... load the SOURCE transcript for the first Clip Transcript
                # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                transcript = Transcript.Transcript(clip.transcripts[0].source_transcript, skipText=True)
            # If any of the transcripts are orphans ...
            else:
##                # Get the list of possible replacement source transcripts
##                transcriptList = DBInterface.list_transcripts(episode.series_id, episode.id)
##                # If only 1 transcript is in the list ...
##                if len(transcriptList) == 1:
##                    # ... use that.
##                    transcript = Transcript.Transcript(transcriptList[0][0])
##                # If there are NO transcripts (perhaps because the Episode is gone, perhaps because it has no Transcripts) ...
##                elif len(transcriptList) == 0:
##                    # ... raise an exception.  We can't locate a transcript if the Episode is gone or if it has no Transcripts.
##                    raise RecordNotFoundError ('Transcript', 0)
##                # If there are multiple transcripts to choose from ...
##                else:
##                    # Initialize a list
##                    strList = []
##                    # Iterate through the list of transcripts ...
##                    for (transcriptNum, transcriptID, episodeNum) in transcriptList:
##                        # ... and extract the Transcript IDs
##                        strList.append(transcriptID)
##                    # Create a dialog where the user can choose one Transcript
##                    dlg = wx.SingleChoiceDialog(self, _('Transana cannot identify the Transcript where this clip originated.\nPlease select the Transcript that was used to create this Clip.'),
##                                                _('Transana Information'), strList, wx.OK | wx.CANCEL)
##                    # Show the dialog.  If the user presses OK ...
##                    if dlg.ShowModal() == wx.ID_OK:
##                        # ... use the selected transcript
##                        transcript = Transcript.Transcript(dlg.GetStringSelection(), episode.number)
##                    # If the user presses Cancel (Esc, etc.) ...
##                    else:
                # ... raise an exception
                raise RecordNotFoundError ('Transcript', 0)
            # Set the active transcript to 0 so the whole interface will be reset
            self.activeTranscript = 0
            # Load the source Transcript
            self.LoadTranscript(episode.series_id, episode.id, transcript.id)
            # Check to see if the load succeeded before continuing!
            if self.currentObj != None:
                # We need to signal that the Visualization needs to be re-drawn.
                self.ChangeVisualization()
                # For each Clip transcript except the first one (which has already been loaded) ...
                for tr in clip.transcripts[1:]:
                    # ... load the source Transcript as an Additional Transcript
                    self.OpenAdditionalTranscript(tr.source_transcript)
                # Mark the Clip as the current selection.  (This needs to be done AFTER all transcripts have been opened.)
                self.SetVideoSelection(clip.clip_start, clip.clip_stop)

                # We need the screen to update here, before the next step.
                wx.Yield()
                # Now let's go through each Transcript Window ...
                for trWin in self.TranscriptWindow.nb.GetCurrentPage().GetChildren():
                    # ... move the cursor to the TRANSCRIPT's Start Time (not the Clip's)
                    trWin.editor.scroll_to_time(clip.transcripts[trWin.panelNum].clip_start + 10)
                    # .. and select to the TRANSCRIPT's End Time (not the Clip's)
                    trWin.editor.select_find(str(clip.transcripts[trWin.panelNum].clip_stop))
                    # update the selection text
                    wx.CallLater(50, trWin.editor.ShowCurrentSelection)
        except:
            (exctype, excvalue, traceback) = sys.exc_info()

            if DEBUG:
                print exctype, excvalue
                import traceback
                traceback.print_exc(file=sys.stdout)
            
            if len(clip.transcripts) == 1:
                msg = _('The Transcript this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
            else:
                msg = _('One of the Transcripts this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
            self.ClearAllWindows(clearAllPanes=True)
            dlg = Dialogs.ErrorDialog(None, msg)
            result = dlg.ShowModal()
            dlg.Destroy()

    def BringTranscriptToFront(self):
        """ The Document Window can contain MANY document tabs, but only one Transcript tab.  Find the Transcript and bring
            it to the front. """
        # Pass this along to the TranscriptWindow
        self.TranscriptWindow.BringTranscriptToFront()
        # Update the Current Object
        if isinstance(self.TranscriptWindow.dlg.editor.TranscriptObj, Transcript.Transcript):
            if self.TranscriptWindow.dlg.editor.TranscriptObj.clip_num == 0:
                self.currentObj = Episode.Episode(self.TranscriptWindow.dlg.editor.TranscriptObj.episode_num)
            else:
                # If a Clip that is open on this computer is being deleted on another computer, an RecordNotFoundError exception
                # will be raised here.
                try:
                    self.currentObj = Clip.Clip(self.TranscriptWindow.dlg.editor.TranscriptObj.clip_num)
                except TransanaExceptions.RecordNotFoundError, e:
                    self.currentObj = None
            if self.currentObj != None:
                # Signal to update Transana's GUI based on this Notebook change
                self.UpdateCurrentObject(self.currentObj)

        else:
            self.currentObj = None

    def GetCurrentDocumentObject(self):
        """ Get the object underlying the currently-open tab in the Document Window """
        return self.TranscriptWindow.dlg.editor.TranscriptObj

    def GetOpenDocumentObject(self, docType, docNum):
        """ if the Document indicated by docNum is currently open, return a pointer to that existing Document
            object.  OBJECTS OBTAINED THIS WAY SHOULD NOT BE EDITED!! """
        # Set the default result to None, expecting that the object in question will NOT be found.
        result = None
        # For each Notebook Page in the Transcript Window ...
        for page in range(self.TranscriptWindow.nb.GetPageCount()):
            # ... for each Splitter Pane on the Notebook Page ...
            for pane in self.TranscriptWindow.nb.GetPage(page).GetChildren():
                # Get the pane's data object
                dataObj = pane.editor.TranscriptObj
                if isinstance(dataObj, docType) and (dataObj.number == docNum):
                    result = dataObj
                    break
        # Return the result
        return result

    def SelectOpenDocumentTab(self, docType, docNum):
        """ Bring the indicated Document / Transcript to the front of the Document Window display """
        # Create a flag for when we can stop looking
        found = False
        # For each Notebook Page in the Transcript Window ...
        for page in range(self.TranscriptWindow.nb.GetPageCount()):
            # ... for each Splitter Pane on the Notebook Page ...
            for pane in self.TranscriptWindow.nb.GetPage(page).GetChildren():
                # Get the pane's data object
                dataObj = pane.editor.TranscriptObj
                # If this is the data object we are looking for ...
                if isinstance(dataObj, docType) and (dataObj.number == docNum):
                    # ... select the appropriate notebook tab ...
                    self.TranscriptWindow.nb.SetSelection(page)
                    # .. and the correct Splitter pane ...
                    pane.ActivatePanel()
                    # ... signal that we are done ...
                    found = True
                    # ... and stop looking at Panes
                    break
            # If we found what we are looking for, we can stop looking at Notebook Tabs too.
            if found:
                break

    def CloseOpenTranscriptWindowObject(self, docType, docNum):
        """ If the Document/Transcript/Quote indicated by docNum is currently open, close it.
            This is used as part of DELETING an object, and adjustments to child objects are
            also made here. """
        # For each Notebook Page in the Transcript Window ... (Count DOWN to avoid an error when a Notebook page is closed!
        for page in range(self.TranscriptWindow.nb.GetPageCount() - 1, -1, -1):

            # If we're dealing with a Clip ... (Clips are different than other types of objects!)
            if (docType == Clip.Clip):
                # For Clips, we ALWAYS close the whole Notebook Page, even if there are multiple transcripts.
                #
                # If we have a Transcript ...
                # Get the FIRST pane's data object (as they will all have the same clip_num!)
                dataObj = self.TranscriptWindow.nb.GetPage(page).GetChildren()[0].editor.TranscriptObj
                if isinstance(dataObj, Transcript.Transcript):
                    # If the Clip Transcript's Clip Number matches the docNum ...
                    if dataObj.clip_num == docNum:

                        if self.TranscriptWindow.nb.GetPageCount() > 1:

                            self.BringTranscriptToFront()

                            # Unload the Media File
                            self.ClearAllWindows(clearAllPanes=True)
                            # ... we can delete the Notebook Page
#                            self.TranscriptWindow.nb.DeleteNotebookPage(page)
                        else:
                            # ... we need to clear the Pane rather than delete it.
                            self.ClearAllWindows(True)
                        
            # If we're dealing with any other object type ...
            else:
                # ... for each Splitter Pane on the Notebook Page ...
                for pane in self.TranscriptWindow.nb.GetPage(page).GetChildren():
                    # Get the pane's data object
                    dataObj = pane.editor.TranscriptObj

                    # If we are deleting a Document, we need to remove the source_document_num from any open Quotes
                    # taken from that document.  (If the object was subsequently edited, the source_document_num could
                    # be incorrectly restored otherwise.)
                    #
                    # Detect if we have a Source Document and a Quote Object
                    if (docType == Document.Document) and isinstance(dataObj, Quote.Quote):
                        # If this Quote is taken from THIS Document ...
                        if dataObj.source_document_num == docNum:
                            # ... clear the source document number to orphan the Quote
                            dataObj.source_document_num = 0

                    # If we have the correct Object Type and the correct Object Number ...
                    if isinstance(dataObj, docType) and (dataObj.number == docNum):
                        # If the Document has been edited ...
                        if pane.editor.modified:
                            # Bring the Document to the front.  (SaveTranscript only works on the current Page)
                            self.TranscriptWindow.nb.SetSelection(page)
                            # Save it (with Prompting!)
                            self.SaveTranscript(1, transcriptToSave=pane.panelNum)
                        # If there's more than one Splitter Pane open ...
                        if len(self.TranscriptWindow.nb.GetPage(page).GetChildren()) > 1:
                            # Clear the Splitter Pane
                            self.ClearAllWindows()
                            # ... we can just delete the Splitter Pane
                            # REDUNDANT!  pane.parent.DeletePanel(pane.panelNum)
                        # Otherwise, if there's more than one Notebook Page open ...
                        elif self.TranscriptWindow.nb.GetPageCount() > 1:
                            # Clear the Notebook Page
#                            self.ClearAllWindows()
                            # ... we can delete the Notebook Page
                            self.TranscriptWindow.nb.DeleteNotebookPage(page)
                        # If there's only one Notebook Page and it has only one Splitter Pane ...
                        else:
                            # ... we need to clear the Pane rather than delete it.
                            self.ClearAllWindows(True)
                        # If we've just closed a Transcript ...
                        if docType == Transcript.Transcript:
                            # ... reset the ControlObject TranscriptNum dictionary
                            self.TranscriptNum = {}

    def UpdateCurrentObject(self, currentObj):
        """ This should be called any time the "current" object is changes, such as when the Transcript Notebook Page
            is changed. """
        # Start exception handling to catch when the current object has been deleted by another user
        try:
            # If the current Object is a Document ...
            if isinstance(currentObj, Document.Document):
                # ... load the Library ...
                tmpLibrary = Library.Library(currentObj.library_num)
                # ... set the Transcript Window title accordingly ...
                self.TranscriptWindow.SetTitle(_('Document').decode('utf8') + u' - ' + tmpLibrary.id + u' > ' + currentObj.id)
                # ... and set the object to work with to the original Document
                tmpCurrentObj = currentObj
            # If we have a Episode Object ...
            elif isinstance(currentObj, Episode.Episode):
                # We already have an episode, so use that.
                tmpEpisode = currentObj
                # ... load the Library ...
                tmpLibrary = Library.Library(tmpEpisode.series_num)
                # ... and we'll work with the Episode
                tmpCurrentObj = currentObj
            # If we have a Transcript Object ...  (SHOULD THIS EVEN HAPPEN??)
            elif (isinstance(currentObj, Transcript.Transcript) and (currentObj.clip_num == 0)):
                # ... load the Episode ...
                tmpEpisode = Episode.Episode(currentObj.episode_num)
                # ... load the Library ...
                tmpLibrary = Library.Library(tmpEpisode.series_num)
                # ... and set the object to work with to the EPISODE
                tmpCurrentObj = tmpEpisode
                # ... set the Transcript Window title accordingly ...
                prompt = unicode('%s - %s > %s > %s', 'utf8')
                self.TranscriptWindow.SetTitle(prompt % (_('Transcript').decode('utf8'), tmpLibrary.id, tmpEpisode.id, currentObj.id))
            elif isinstance(currentObj, Quote.Quote):
                # ... set the Transcript Window title accordingly ...
                self.TranscriptWindow.SetTitle(_('Quote').decode('utf8') + u' - ' + currentObj.GetNodeString(True))
                # ... and set the object to work with to the original Document
                tmpCurrentObj = currentObj
            elif isinstance(currentObj, Clip.Clip):
                # ... set the Transcript Window title accordingly ...
                self.TranscriptWindow.SetTitle(_('Clip').decode('utf8') + u' - ' + currentObj.GetNodeString(True))
                # ... and set the object to work with to the original Clip
                tmpCurrentObj = currentObj
            elif isinstance(currentObj, Transcript.Transcript) and (currentObj.clip_num > 0):
                try:
                    # ... load the Clip ...
                    tmpClip = Clip.Clip(currentObj.clip_num)
                    # ... and set the object to work with to the CLIP
                    tmpCurrentObj = tmpClip
                    # ... set the Transcript Window title accordingly ...
                    self.TranscriptWindow.SetTitle(_('Clip').decode('utf8') + u' - ' + tmpClip.GetNodeString(True))
                except TransanaExceptions.RecordNotFoundError, e:
                    # If the record is not found, that's because ANOTHER USER has deleted it!!
                    # We have to fake it here!
                    tmpCurrentObj = Clip.Clip()
                    tmpCurrentObj.number = currentObj.clip_num
                    
            else:
                tmpCurrentObj = currentObj
                
                print "ControlObjectClass.UpdateCurrentObject():", self.currentObj == currentObj
                print type(self.currentObj), type(currentObj)
                print currentObj
                print
        # If the current object is not found, 
        except TransanaExceptions.RecordNotFoundError, e:
            # The current object has been deleted, so signal that!
            tmpCurrentObj = None

        
        # Remove any tabs in the Data Window beyond the Database Tab
        self.DataWindow.DeleteTabs()
        # This method can have problems during MU object deletion.  Therefore, start exception handling.
        try:
            if tmpCurrentObj == None:
                pass

            elif isinstance(tmpCurrentObj, Document.Document):
                tmpLibrary = Library.Library(tmpCurrentObj.library_num)
                # Add the Document Quotes Tab to the DataWindow
                self.DataWindow.AddItemsTab(libraryObj=tmpLibrary, dataObj=tmpCurrentObj)
                # Add the Selected Document Quotes Tab to the DataWindow
                self.DataWindow.AddSelectedItemsTab(libraryObj=tmpLibrary, dataObj=tmpCurrentObj,
                                                    textPos=self.TranscriptWindow.dlg.editor.GetCurrentPos(),
                                                    textSel=self.TranscriptWindow.dlg.editor.GetSelection())
                # Add the Keyword Tab to the DataWindow
                self.DataWindow.AddKeywordsTab(seriesObj = tmpLibrary, documentObj = tmpCurrentObj)
                # Set the Visualization Window's Visualization Object
                self.VisualizationWindow.SetVisualizationObject(tmpCurrentObj)

            elif isinstance(tmpCurrentObj, Episode.Episode):
                # Add the Episode Clips Tab to the DataWindow
                self.DataWindow.AddItemsTab(libraryObj = tmpLibrary, dataObj = tmpEpisode)
                # Add the Selected Episode Clips Tab to the DataWindow
                self.DataWindow.AddSelectedItemsTab(libraryObj = tmpLibrary, dataObj = tmpEpisode, timeCode=self.GetVideoPosition())
                # Add the Keyword Tab to the DataWindow
                self.DataWindow.AddKeywordsTab(seriesObj = tmpLibrary, episodeObj = tmpEpisode)

                # For reasons that elude me, "if self.currentObj != tmpCurrentObj" didn't work, but this does!                
                # If we're changing the underlying object ...
                if (type(self.currentObj) != type(tmpCurrentObj)) or (self.currentObj.number != tmpCurrentObj.number):
                    # Set the Visualization Window's Visualization Object
                    self.VisualizationWindow.SetVisualizationObject(tmpEpisode)

            elif isinstance(tmpCurrentObj, Quote.Quote):
                tmpCollection = Collection.Collection(tmpCurrentObj.collection_num)
                # Add the Keyword Tab to the DataWindow
                self.DataWindow.AddKeywordsTab(collectionObj = tmpCollection, quoteObj = tmpCurrentObj)
                # Set the Visualization Window's Visualization Object
                self.VisualizationWindow.SetVisualizationObject(tmpCurrentObj)

            elif isinstance(tmpCurrentObj, Clip.Clip):
                tmpCollection = Collection.Collection(tmpCurrentObj.collection_num)
                # Add the Keyword Tab to the DataWindow
                self.DataWindow.AddKeywordsTab(collectionObj = tmpCollection, clipObj = tmpCurrentObj)
                # For reasons that elude me, "if self.currentObj != tmpCurrentObj" didn't work, but this does!                
                # If we're changing the underlying object ...
                # Add check of VisualizationWindow's current type because if you hit PLAY with a Document or Quote
                # selected and a Clip in the background, the visualization wasn't updating correctly!!
                if (type(self.currentObj) != type(tmpCurrentObj)) or (self.currentObj.number != tmpCurrentObj.number) or \
                   (type(self.currentObj) != (type(self.VisualizationWindow.VisualizationObject))):
                    # Set the Visualization Window's Visualization Object
                    self.VisualizationWindow.SetVisualizationObject(tmpCurrentObj)

            else:

                print "ControlObjectClass.UpdateCurrentObject():  %s not implemented." % type(tmpCurrentObj)

        # If we can't find a record to load the appropriate object, it's probably been deleted by another user.
        except TransanaExceptions.RecordNotFoundError, e:

#            print "ControlObjectClass.UpdateCurrentObject(): ", type(tmpCurrentObj), tmpCurrentObj.id
#            print sys.exc_info()[0]
#            print sys.exc_info()[1]
            
            pass

        # Set the new Current Object
        self.currentObj = tmpCurrentObj

        # if we are working with an Episode, this may be switching between multiple transcipts.
        if isinstance(self.currentObj, Episode.Episode):
            # Update the Selection Text
            self.UpdateSelectionTextLater()

    def RemoveQuoteFromOpenDocument(self, quote_num, doc_num):
        """ If a Quote is deleted and the source document for that quote is currently open, we need to remove
            the quote_dict reference to that quote.  We pass source document number and quote number, as the
            quote itself may have already been deleted. """
        # If the quote has a known Source Document ...
        if doc_num > 0:
            # ... see if the Document is currently open and get it
            docObj = self.GetOpenDocumentObject(Document.Document, doc_num)
            # If the Document IS currently open ...
            if docObj != None:
                # ... remove the specified Quote from the Document
                del(docObj.quote_dict[quote_num])

    def LoadSnapshot(self, snapshot):
        """ Load the SnapshotWindow for the Snapshot object passed in """
        # Assume no Snapshot Window exists for this snapshot
        windowOpen = False
        # If we have a known Snapshot Number ...
        if snapshot.number != 0:
            # ... iterate through all Snapshot Windows ...
            for snapshotWindow in self.SnapshotWindows:
                # ... see if there is already a Snapshot Window open for the specified Snapshot
                if snapshotWindow.obj.number == snapshot.number:
                    # If so, show the windows ...
                    snapshotWindow.Show()
                    # ... raise it to the top of the stack ...
                    snapshotWindow.Raise()
                    # ... and if it's Iconized (minimized) ...
                    if snapshotWindow.IsIconized():
                        # ... then un-minimize it.  (This is different from maximizing it, of course!)
                        snapshotWindow.Iconize(False)
                    # Note that an open windows was found
                    windowOpen = True
                    # We can stop looking once we've found an open windows
                    break
        # If we did NOT find an open window ...
        if not windowOpen:
            # Start Exception Handling
            try:
                title = _("Snapshot").decode('utf8') + u' - ' + snapshot.GetNodeString(True)
                # ... create a new Snapshot Window ...
                snapshotDlg = SnapshotWindow.SnapshotWindow(self.MenuWindow, -1, title, snapshot)
                # ... and add this to the list of open Snapshot Windows
                self.SnapshotWindows.append(snapshotDlg)
                # Iterate through the existing Snapshot Windows
                for win in self.SnapshotWindows:
                    # For all windows except the newest one ...
                    if win != snapshotDlg:
                        # ... add the looping window's ID to the Window Menu
                        snapshotDlg.AddWindowMenuItem(win.obj.id, win.obj.number)
                    # Add the new window's ID to the looping window's Window menu
                    win.AddWindowMenuItem(snapshot.id, snapshot.number)
                # Add this to the Menu Window's Window's menu
                self.MenuWindow.AddWindowMenuItem(snapshot.id, snapshot.number)
                    
            # If an ImageLoadError occurs ...
            except TransanaExceptions.ImageLoadError, exception:
                # ... report the error to the user
                dlg = Dialogs.ErrorDialog(self.MenuWindow, exception.explanation)
                dlg.ShowModal()
                dlg.Destroy()

            # Get the current selection(s) from the Database Tree
            selItems = self.DataWindow.DBTab.tree.GetSelections()
            # If there are one or more items selected ...
            if len(selItems) >= 1:
                # ... get the item data from the first selection
                selData = self.DataWindow.DBTab.tree.GetPyData(selItems[0])
            # If NO items are selected ...
            else:
                # ... then there's no item data to get
                selData = None

            # If no items are selected or the item selected is NOT a Search Collection or Search Clip ...
            if (selData == None) or not (selData.nodetype in ['SearchCollectionNode', 'SearchClipNode']):
                # Let's make sure this clip is displayed in the Database Tree
                nodeList = (_('Collections'),) + snapshot.GetNodeData()
                # Now point the DBTree (the notebook's parent window's DBTab's tree) to the loaded Clip
                self.DataWindow.DBTab.tree.select_Node(nodeList, 'SnapshotNode')

    def SelectSnapshotWindow(self, itemName, itemNumber, selectInDataWindow=False):
        """ Select the indicated Snapshot Window """
        # Assume the window will NOT be found
        winFound = False
        # Iterate through the Snapshot Windows
        for win in self.SnapshotWindows:
            # If we have the correct window ...
            if (itemName == win.obj.id) and (itemNumber == win.obj.number):
                # ... if we're supposed to select the Snapshot in the Data Window ...
                if selectInDataWindow:
                    # ... bet the Snapshot's Node Data ...
                    nodeList = (_('Collections'),) + win.obj.GetNodeData()
                    # ... and point the DBTree to the loaded Snapshot
                    self.DataWindow.DBTab.tree.select_Node(nodeList, 'SnapshotNode')
                # ... bring it to the top of the window stack ...
                win.Raise()
                # ... make sure it's not minimized ...
                win.Iconize(False)
                # ... give it focus ...
                win.SetFocus()
                # Note that the window has been found
                winFound = True
                # ... and stop looking
                break

        # Return an indicator of whether the window was found
        return winFound

    def RemoveSnapshotWindow(self, itemName, itemNumber):
        """ Remove a Snapshot Window and all references from Transana's Interface """
        # Iterate through the Snapshot Windows.
        for snapshotWindow in self.SnapshotWindows:
            # When we find the Snapshot Window reference for this Snapshot ...
            if itemNumber == snapshotWindow.obj.number:
                # ... remove it from the List of Snapshot Windows
                self.SnapshotWindows.remove(snapshotWindow)
        # Iterate through the Snapshot Windows again.  (remove() above changed the number of items, so skips one!)
        for snapshotWindow in self.SnapshotWindows:
            # Remove this item from the Snapshot Windows' Window menu
            snapshotWindow.DeleteWindowMenuItem(itemName, itemNumber)
        # Remove this from the Menu Window's Window's menu
        self.MenuWindow.DeleteWindowMenuItem(itemName, itemNumber)

    def GetOpenSnapshotWindows(self, editableOnly=False):
        """ Return a list of all Snapshot Windows that are open """
        # Initialize values to be returned
        values = []
        # Iterate through the Snapshot Windows
        for win in self.SnapshotWindows:
            # If we want all windows, or if the current window is editable ...
            if (not editableOnly) or win.editTool.IsToggled():
                # ... then add the window to the return values
                values.append(win)
        # Return the Return values
        return values

    def UpdateWindowMenu(self, oldSnapshotName, oldSnapshotNumber, newSnapshotName, newSnapshotNumber):
        """ Update all Window Menus when a Snapshot Window changes Snapshots through Prev / Next buttons """
        # See if the NEW window is already open somewhere
        if self.SelectSnapshotWindow(newSnapshotName, newSnapshotNumber):
            # If so, iterate through the open Snapshot Windows ...
            for win in self.SnapshotWindows:
                # If this Snapshot Window matches the one we're trying to open ...
                if (win.obj.id == newSnapshotName) and (win.obj.number == newSnapshotNumber):
                    # ... then close it!  (This will save unsaved edits!
                    win.Close()
        # Iterate through the open Snapshot Windows ...
        for win in self.SnapshotWindows:
            # ... and trigger an update to their MenuWindows
            win.UpdateWindowMenuItem(oldSnapshotName, oldSnapshotNumber, newSnapshotName, newSnapshotNumber)
        # Update the MenuWindow's Windows Menu
        self.MenuWindow.UpdateWindowMenuItem(oldSnapshotName, oldSnapshotNumber, newSnapshotName, newSnapshotNumber)

    def ShowNotesBrowser(self):
        """ Bring the Notes Browser to the front """
        # Raise the Notes Browser Window to the top
        self.NotesBrowserWindow.Raise()
        # If the Notes Browser is minimized ...
        if self.NotesBrowserWindow.IsIconized():
            # ... restore it to full size
            self.NotesBrowserWindow.Iconize(False)
        # Push the Visualization Window behind it
        self.VisualizationWindow.Lower()
        # Push the Video Window behind it
        self.VideoWindow.Lower()
        # Push the Transcript Window behind the Notes Browser
        self.TranscriptWindow.Lower()
        # Push the Data Window behind it
        self.DataWindow.Lower()
        # For each Snapshot Window ...
        for win in self.SnapshotWindows:
            # ... push the Snapshot behind the Notes Browser
            win.Lower()
        # For each Report Window ...
        for win in self.ReportWindows.keys():
            # ... push the Report behind the Notes Browser
            self.ReportWindows[win].Lower()

    def AddReportWindow(self, reportWindow):
        """ Add a Report Window to Transana's Interface """
        # Increment the unique Report Number
        self.reportNumber += 1
        # Remember the report in the ReportWindows dictionary
        self.ReportWindows[self.reportNumber] = reportWindow
        # Add the unique Report Number to the report
        reportWindow.reportNumber = self.reportNumber
        # Add the Report to the Window Menu
        self.MenuWindow.AddWindowMenuItem(reportWindow.title, self.reportNumber)

    def SelectReportWindow(self, reportName, reportNumber):
        """ Select the indicated Report Window """
        # If the desired report number is in the ReportWindows keys ...
        if reportNumber in self.ReportWindows.keys():
            # ... select the correct window
            win = self.ReportWindows[reportNumber]
            # ... bring it to the top of the window stack ...
            win.Raise()
            # ... make sure it's not minimized ...
            win.Iconize(False)
            # ... give it focus ...
            win.SetFocus()

    def RemoveReportWindow(self, reportName, reportNumber):
        """ Remove a Report Window from Transana's Interface """
        # If the report window still exists ... (sometimes a close event will get double-called and the report will already be gone!)
        if self.ReportWindows.has_key(reportNumber):
            # Delete the Report Window from the ReportWindows dictionary
            del(self.ReportWindows[reportNumber])
            # Remove this from the Menu Window's Window's menu
            self.MenuWindow.DeleteWindowMenuItem(reportName, reportNumber)

    def OpenAdditionalDocument(self, documentNum, libraryID='', isDocument=True):
        """ Open an additional document without replacing the current one """

        if DEBUG:
            print "ControlObjectClass.OpenAdditionalDocument():"
            print '  ', self.TranscriptWindow.nb.GetPageCount(), self.TranscriptWindow.nb.GetSelection()
            print 'ControlObject.CurrentObj:'
            print '  ', self.currentObj
            print 'editor.TranscriptObj:'
            print '  ', self.TranscriptWindow.dlg.editor.TranscriptObj
            print

        # First, let's see if the current Notebook Page houses a Document
        if isinstance(self.TranscriptWindow.nb.GetPage(self.TranscriptWindow.nb.GetSelection()).GetChildren()[0].editor.TranscriptObj, Document.Document):
            # Add a Pane to the TranscriptWindow's Current Notebook Page
            self.TranscriptWindow.AddPanel()
            # Select the newly created Panel to activate it
            self.TranscriptWindow.nb.GetCurrentPage().activePanel = len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) - 1

            # Get out Document object from the database
            documentObj = Document.Document(documentNum)
            # Load the desired Transcript into the new Pane
            self.TranscriptWindow.LoadTranscript(documentObj)

            # Set the Notebook's Page Text too
            self.TranscriptWindow.nb.SetPageText(self.TranscriptWindow.nb.GetSelection(), _('Multiple Documents'))

            # Update display for this object (e.g. Data Window tabs)
            self.UpdateCurrentObject(documentObj)
            
        # Set focus to the new Transcript's Editor (so that CommonKeys work on the Mac)
        self.TranscriptWindow.dlg.editor.SetFocus()

    def OpenAdditionalTranscript(self, transcriptNum, seriesID='', episodeID='', isEpisodeTranscript=True):
        """ Open an additional Transcript without replacing the current one """

        if DEBUG:
            print "ControlObjectClass.OpenAdditionalTranscript():"
            print '  ', self.TranscriptWindow.nb.GetPageCount(), self.TranscriptWindow.nb.GetSelection()
            print 'ControlObject.CurrentObj:'
            print '  ', self.currentObj
            print 'editor.TranscriptObj:'
            print '  ', self.TranscriptWindow.dlg.editor.TranscriptObj
            print

        # First, let's see if there's already a video loaded in the system.  Iterate through all Notebook Pages.
        for y in range(self.TranscriptWindow.nb.GetPageCount()):
            # If the page's first Splitter pane's editor's TranscriptObj is a Transcript ...
            if isinstance(self.TranscriptWindow.nb.GetPage(y).GetChildren()[0].editor.TranscriptObj, Transcript.Transcript):
                # ... then select this as the active Notebook Page, effectively limiting Transana to ONE Episode at a time.
                self.TranscriptWindow.nb.SetSelection(y)
                # We don't need to keep looking.
                break

        # Add a Pane to the TranscriptWindow's Current Notebook Page
        self.TranscriptWindow.AddPanel()
        # Select the newly created Panel to activate it
        self.TranscriptWindow.nb.GetCurrentPage().activePanel = len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) - 1
        
##        # Create a new Transcript Window
##        newTranscriptWindow = TranscriptionUI.TranscriptionUI(TransanaGlobal.menuWindow, includeClose=True)
##        # Register this new Transcript Window with the Control Object (self)
##        self.Register(Transcript=newTranscriptWindow)
##        # Register the Control Object (self) with the new Transcript Window
##        newTranscriptWindow.Register(self)
        
        # Get out Transcript object from the database
        transcriptObj = Transcript.Transcript(transcriptNum)
        # Load the desired Transcript into the new Pane
        self.TranscriptWindow.LoadTranscript(transcriptObj)

        self.UpdateCurrentObject(transcriptObj)

        # If we have an Episode Transcript, it needs a Window title.  (Clip titles are handled in the calling routine.)
        if isEpisodeTranscript:
            # If we haven't been sent an Episode ID ...
            if episodeID == '':
                # ... get the Episode data based on the Transcript Object ...
                episodeObj = Episode.Episode(transcriptObj.episode_num)
                # ... and note the Episode ID
                episodeID = episodeObj.id
            # Create the Notebook Page prompt
            prompt = unicode(_('Transcripts for %s'), 'utf8')
            # Set the Notebook's Page Text too
            self.TranscriptWindow.nb.SetPageText(self.TranscriptWindow.nb.GetSelection(), prompt % episodeID)
            
        # Add the new transcript's number to the list that tracks the numbers of the open transcripts.
        self.TranscriptNum[transcriptObj.number] = (self.TranscriptWindow.nb.GetSelection(), self.TranscriptWindow.nb.GetPage(self.TranscriptWindow.nb.GetSelection()).activePanel)

        # Position the Cursor / Highlight in the new Transcript
        self.TranscriptWindow.UpdatePosition(self.VideoWindow.GetCurrentVideoPosition())

        # Enable the Multiple Transcript buttons
        self.TranscriptWindow.UpdateMultiTranscriptButtons(True)

        # Set focus to the new Transcript's Editor (so that CommonKeys work on the Mac)
        self.TranscriptWindow.dlg.editor.SetFocus()

##        if DEBUG:
##            print "ControlObjectClass.OpenAdditionalTranscript(%d)  %d" % (transcriptNum, self.activeTranscript)
##            for x in range(len(self.TranscriptWindow)):
##                print x, self.TranscriptWindow[x].transcriptWindowNumber, self.TranscriptNum[x]
##            print
        
    def CloseAdditionalTranscript(self, transcriptNum):
        """ Close a secondary transcript """

        print "ControlObjectClass.CloseAdditionalTranscript():  Is this used??  11-14-2014"
        
        # If we're closeing a transcript other than the active transcript ...
        if self.activeTranscript != transcriptNum and not self.shuttingDown:
            # ... remember which transcript WAS active ...
            prevActiveTranscript = self.activeTranscript
            # ... and make the one we're supposed to close active.
            self.activeTranscript = transcriptNum
        # If we're closing the active transcript ...
        else:
            # ... then focus should switch to the top transcript, # 0
            prevActiveTranscript = 0
        # If the prevActiveTranscript is about to be closed, we need to reduce it by one to avoid
        # problems on the Mac.
        if prevActiveTranscript == len(self.TranscriptWindow) - 1:
            prevActiveTranscript = self.activeTranscript - 1
        # Before we do anything else, let's save the current transcript if it's been modified.
        if self.TranscriptWindow.TranscriptModified():
            if TransanaConstants.partialTranscriptEdit:
                self.SaveTranscript(1, cleardoc=1, continueEditing=False)
            else:
                self.SaveTranscript(1, cleardoc=1)
        if transcriptNum == 0:
            (left, top) = self.TranscriptWindow.GetPositionTuple()
            self.TranscriptWindow.SetPosition(wx.Point(left, top))
        # Set the frame focus to the Previous active transcript (I'm not convinced this does anything!)
        self.TranscriptWindow.dlg.SetFocus()
        # Update the Active Transcript number
        self.activeTranscript = prevActiveTranscript
        # If there's only one transcript left ...
        if len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) == 1:
            # ... Disable the Multiple Transcript buttons
            self.TranscriptWindow.UpdateMultiTranscriptButtons(False)

    def SaveAllTranscriptCursors(self):
        """ Save the current cursor position or selection for all open Transcript windows """
        # For each Transcript Window Notebook Page ...
        for pageNum in range(self.TranscriptWindow.nb.GetPageCount()):
            # ... for each Splitter Panel on each Page ...
            for pane in self.TranscriptWindow.nb.GetPage(pageNum).GetChildren():
                # ... save the cursorPosition
                pane.editor.SaveCursor()

    def RestoreAllTranscriptCursors(self):
        """ Restore the previously saved cursor position or selection for all open Transcript windows """
        # For each Transcript Window Notebook Page ...
        for pageNum in range(self.TranscriptWindow.nb.GetPageCount()):
            # ... for each Splitter Panel on each Page ...
            for pane in self.TranscriptWindow.nb.GetPage(pageNum).GetChildren():
                # ... if the editor HAS a saved cursorPosition ...
                if pane.editor.cursorPosition != 0:
                    # ... restore the cursorPosition or selection to the Editor
                    pane.editor.RestoreCursor()

    def ClearAllWindows(self, clearAllTabs=False, clearAllPanes=False):
        """ Clears all windows and resets all objects related to Media Files.
            If clearAllTabs is True, we close all Documents and Quotes too! """
        # Let's stop the media from playing
        self.VideoWindow.Stop()
        
        # If we're clearing ALL tabs ...
        if clearAllTabs:
            # ... we want a list of all Notebook Page Numbers in descending order
            tabNumList = range(self.TranscriptWindow.nb.GetPageCount() - 1, -1, -1)
            # If we're clearing all tabs, we want to clear everything.
            closeEverything = True
        # If we're NOT clear ALL tabs ...
        else:
            # ... we want a list with just the current Notebook Page Number in it
            tabNumList = [self.TranscriptWindow.nb.GetSelection()]
            # We don't actually know if we want to clear everything yet!  False for now...
            closeEverything = False

        # Save and unlock the Document Records, if any are locked
        # For each Notebook Tab ... (Count backwards from the highest number, as provided in tabNumList!)
        for tabNum in tabNumList:
            # Select the Notebook Page
            self.TranscriptWindow.nb.SetSelection(tabNum)

            if clearAllTabs:
                # While there are additional Transcript windows open ...
                for paneNum in range(len(self.TranscriptWindow.nb.GetPage(tabNum).GetChildren()) - 1, -1, -1):
                    self.TranscriptWindow.nb.GetPage(tabNum).ActivatePanel(paneNum)
                    # Save the transcript
                    if TransanaConstants.partialTranscriptEdit:
                        self.SaveTranscript(1, transcriptToSave=len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) - 1, continueEditing=False)
                    else:
                        self.SaveTranscript(1, transcriptToSave=len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) - 1)
                    
                    # Clear Document Window
                    self.TranscriptWindow.CloseCurrent(None)

            else:
                if clearAllPanes:
                    panes = range(len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) - 1, -1, -1)
                else:
                    panes = (self.TranscriptWindow.nb.GetCurrentPage().activePanel,)
                for pane in panes:
                    self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(pane)
                    # Save the transcript
                    if TransanaConstants.partialTranscriptEdit:
                        self.SaveTranscript(1, transcriptToSave=pane, continueEditing=False)
                    else:
                        self.SaveTranscript(1, transcriptToSave=pane)

                    # If there's only one PANE ...
                    if len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) == 1:
                        # ... then we are in effect closing everything.
                        closeEverything = True

                    self.TranscriptWindow.CloseCurrent(None)

        if len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) > 1:
            str = _('Multiple Documents')
            self.TranscriptWindow.nb.SetPageText(self.TranscriptWindow.nb.GetSelection(), str)
            # Identify the loaded Object
            str = _('Document')
            self.TranscriptWindow.SetTitle(str)
        else:
            if self.TranscriptWindow.dlg.editor.TranscriptObj != None:
                str = self.TranscriptWindow.dlg.editor.TranscriptObj.id
            else:
                str = _('(No Document Loaded)')
            self.TranscriptWindow.nb.SetPageText(self.TranscriptWindow.nb.GetSelection(), str)
            # Identify the loaded Object
            str = _('Document')
            self.TranscriptWindow.SetTitle(str)

        if closeEverything:
            # Clear the currently loaded object, as there is none
            self.currentObj = None

            # Reset the ControlObject's TranscriptNum dictionary
            self.TranscriptNum = {}

            # Clear the Menu Window (Reset menus to initial state)
            self.MenuWindow.ClearMenus()
            # Clear Visualization Window
            self.VisualizationWindow.ClearVisualization()
            # Clear the Video Window
            self.ClearMediaWindow()

            # Clear the Data Window
            self.DataWindow.ClearData()

        # Force the screen updates
        # there can be an issue with recursive calls to wxYield, so trap the exception ...
        try:
            wx.Yield()
        # ... and ignore it!
        except:
            pass

    def GetNewDatabase(self):
        """ Close the old database and open a new one. """
        # set the active transcript to 0 so multiple transcript will be cleared
        self.activeTranscript = 0
        # Clear all existing Data
        self.ClearAllWindows(clearAllTabs=True)
        # Close all Snapshot Windows
        self.CloseAllImages()
        # Close all Reports
        self.CloseAllReports()
        # If we're in multi-user ...
        if not TransanaConstants.singleUserVersion:
            # ... stop the Connection Timer so it won't fire while the Database is closed
            TransanaGlobal.connectionTimer.Stop()
        # Close the existing database connection
        DBInterface.close_db()
        # Reset the global encoding to UTF-8 if the Database supports it
        if (TransanaGlobal.DBVersion >= u'4.1') or \
           (not TransanaConstants.DBInstalled in ['MySQLdb-embedded', 'MySQLdb-server', 'PyMySQL']):
            TransanaGlobal.encoding = 'utf8'
        # Otherwise, if we're in Russian, change the encoding to KOI8r
        elif TransanaGlobal.configData.language == 'ru':
            TransanaGlobal.encoding = 'koi8_r'
        # If we're in Chinese, change the encoding to the appropriate Chinese encoding
        elif TransanaGlobal.configData.language == 'zh':
            TransanaGlobal.encoding = TransanaConstants.chineseEncoding
        # If we're in East Europe Encoding, change the encoding to 'iso8859_2'
        elif TransanaGlobal.configData.language == 'easteurope':
            TransanaGlobal.encoding = 'iso8859_2'
        # If we're in Greek, change the encoding to 'iso8859_7'
        elif TransanaGlobal.configData.language == 'el':
            TransanaGlobal.encoding = 'iso8859_7'
        # If we're in Japanese, change the encoding to cp932
        elif TransanaGlobal.configData.language == 'ja':
            TransanaGlobal.encoding = 'cp932'
        # If we're in Korean, change the encoding to cp949
        elif TransanaGlobal.configData.language == 'ko':
            TransanaGlobal.encoding = 'cp949'
        # Otherwise, fall back to Latin-1
        else:
            TransanaGlobal.encoding = 'latin1'
        # If a new database login fails three times, we need to close the program.
        # Initialize a counter to track that.
        logonCount = 1
        # Flag if Logon succeeds
        loggedOn = False
        # Keep trying for three tries or until successful
        while (logonCount <= 3) and (not loggedOn):
            # Increment logon counter
            logonCount += 1
            # Call up the Username and Password Dialog to get new connection information
            if DBInterface.establish_db_exists():
                # Now update the Data Window
                self.DataWindow.DBTab.tree.refresh_tree()
                # Indicate successful logon
                loggedOn = True
            # If logon fails, inform user and offer to try again twice.
            elif logonCount <= 3:
                # Create a Dialog Box
                dlg = Dialogs.QuestionDialog(self.MenuWindow, _('Transana was unable to connect to the database.\nWould you like to try again?'),
                                         _('Transana Database Connection'))
                # If the user does not want to try again, set the counter to 4, which will cause the program to exit
                if dlg.LocalShowModal() == wx.ID_NO:
                    logonCount = 4
                # Clean up the Dialog Box
                dlg.Destroy()
            # If we're in multi-user and we successfully logged in ...
            if not TransanaConstants.singleUserVersion and loggedOn:
                # ... start the Connection Timer.  This attempts to prevent the "Connection to Database Lost" error by
                # running a very small query every 10 minutes.  See Transana.py.
                TransanaGlobal.connectionTimer.Start(600000)

        # If the Database Connection fails ...
        if not loggedOn:
            # ... Close Transana
            self.MenuWindow.OnFileExit(None)

    def ShowDataTab(self, tabValue):
        """ Changes the visible tab in the notebook in the Data Window """
        if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
            # Display the Keywords Tab
            self.DataWindow.nb.SetSelection(tabValue)
            # Refresh the display for OS X
            self.DataWindow.Refresh()

    def ProcessCommonKeyCommands(self, event):
        """ Process keyboard commands common to several of Transana's main windows """

#        print "ControlObjectClass.ProcessCommonKeyCommands():", event.ShiftDown(), event.ControlDown(), event.CmdDown(), event.AltDown(), event.GetKeyCode()
        
        # Assume the key WILL be processed here in this this method
        keyProcessed = True
        # Extract the key code from the event passed in
        c = event.GetKeyCode()
        # Determine if there are modifiers
        hasMods = event.AltDown() or event.ControlDown() or event.CmdDown() or event.ShiftDown()
        # Note whether there is something loaded in the main interface
        loaded = (self.currentObj != None)
        
        # F1 = Focus on Menu Window
        if (c == wx.WXK_F1) and not hasMods:
            # Set the focus on the Menu Window
            self.MenuWindow.tmpCtrl.SetFocus()

        # F2 = Focus on Visualization Window
        elif (c == wx.WXK_F2) and not hasMods:
            # Set the focus to the Visualization Window
            self.VisualizationWindow.SetFocus()

        # F3 = Focus on Video Window
        elif (c == wx.WXK_F3) and not hasMods:
            # Set the focus to the Video Window
            self.VideoWindow.SetFocus()

        # F4 = Focus on Transcript Window
        elif (c == wx.WXK_F4) and not hasMods:
            # Determine where the focus currently is (I don't exactly understand why this works.  It's something about
            # this being a "static function")
            tmpFocVal = self.TranscriptWindow.dlg.editor.FindFocus()
            # Now set the focus to the currently active transcript window
            self.TranscriptWindow.dlg.editor.SetFocus()
            # If the focus didn't change ...
            if tmpFocVal == self.TranscriptWindow.dlg.editor.FindFocus():
                # ... then the Transcript Window already HAD focus.  So see if there is more than one Transcript Pane on the
                # current Transcript Window Notebook Page after the active Panel ...
                if self.TranscriptWindow.nb.GetCurrentPage().activePanel < len(self.TranscriptWindow.nb.GetCurrentPage().GetChildren()) - 1:
                    # ... move to the next Splitter Pane
                    self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(self.TranscriptWindow.nb.GetCurrentPage().activePanel + 1)
                # If we're at the last Splitter Pane AND there's more than one Notebook Page ...
                elif self.TranscriptWindow.nb.GetPageCount() > 1:
                    # ... if we're not at the last Notebook Page ...
                    if self.TranscriptWindow.nb.GetSelection() < self.TranscriptWindow.nb.GetPageCount() - 1:
                        # ... select the Next Notebook Page ...
                        self.TranscriptWindow.nb.SetSelection(self.TranscriptWindow.nb.GetSelection() + 1)
                        # ... and activate the firstSplitterPane on that Notebook Page
                        self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(0)
                    # If we ARE at the last of multiple Notebook Pages ...
                    else:
                        # ... select the First Notebook Page ...
                        self.TranscriptWindow.nb.SetSelection(0)
                        # ... and activate the first Splitter on that Notebook Page
                        self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(0)
                # If there's only one Notebook Page and we're at its last Splitter Panel ...
                else:
                    # ... activate the first (possibly only) Splitter Pane on that Notebook Page
                    self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(0)
                # Now set the focus to the NEW Transcript notebook page and splitter pane
                self.TranscriptWindow.dlg.editor.SetFocus()
                # Update the Document Window Toolbar
                self.TranscriptWindow.UpdateToolbarAppearance()

        # F5 = Focus on Data Window
        elif (c == wx.WXK_F5) and not hasMods:
            # If the Data Window is currently showing the Database tab ...
            if self.DataWindow.nb.GetPageText(self.DataWindow.nb.GetSelection()) == unicode(_("Database"), 'utf8'):
                # ... set the focus to the Database Tree on that tab
                self.DataWindow.DBTab.tree.SetFocus()
            # Otherwise ...
            else:
                # ... just focus on the Data Window Notebook control.
                self.DataWindow.nb.SetFocus()

        # F6 = Toggle Edit / Read Only Mode
        elif (c == wx.WXK_F6) and not hasMods and loaded:
            # Set the focus to the Transcript
            self.TranscriptWindow.dlg.editor.SetFocus()
            # Toggle the Read Only Button on the Transcript Toolbar
            self.TranscriptWindow.toolbar.ToggleTool(
                self.TranscriptWindow.CMD_READONLY_ID,
                self.TranscriptWindow.dlg.editor.get_read_only())
            # Emulate the Press of the Read Only Button by calling its event directly
            self.TranscriptWindow.OnReadOnlySelect(event)

        # F12 for Windows, Cmd-F12 for Mac are Quick Save
        elif (((not 'wxMac' in wx.PlatformInfo) and (c == wx.WXK_F12) and (not hasMods)) or \
              (('wxMac' in wx.PlatformInfo) and (c == wx.WXK_F12) and (event.CmdDown() and not event.ShiftDown() and not event.AltDown()))) and \
              loaded:
            # if the transcript is in EDIT mode ...
            if not self.TranscriptWindow.dlg.editor.get_read_only():
                # ... save it
                self.TranscriptWindow.dlg.editor.save_transcript()

        # Ctrl-A is Rewind 10 seconds
        elif (c == ord('A')) and event.ControlDown() and loaded:
            # Get the current video position
            vpos = self.GetVideoPosition()
            # Rewind 10 seconds
            self.SetVideoStartPoint(vpos-10000)
            # Explicitly tell Transana to play to the end of the Episode/Clip
            self.SetVideoEndPoint(-1)
            # Play should always be initiated on Ctrl-A, with no auto-rewind
            self.Play(False)

        # Ctrl-D is Stop / Start without Rewind
        elif (c == ord('D')) and event.ControlDown() and loaded:
            if not self.IsPlaying():
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
            # Play/Pause without rewinding
            self.PlayPause(False)

        # Ctrl-F is Fast Forward 10 seconds
        elif (c == ord('F')) and event.ControlDown() and loaded:
            # Get the current video position
            vpos = self.GetVideoPosition()
            # Fast Forward 10 seconds
            self.SetVideoStartPoint(vpos+10000)
            # Explicitly tell Transana to play to the end of the Episode/Clip
            self.SetVideoEndPoint(-1)
            # Play should always be initiated on Ctrl-F, with no auto-rewind
            self.Play(False)

        # Ctrl-P is Play Previous Time-Coded Segment
        elif (c == ord("P")) and event.ControlDown() and loaded:
            # Get the value for the previous time code
            start_timecode = self.TranscriptWindow.dlg.editor.PrevTimeCode()
            # If there WAS a Previous Segment ....
            if (start_timecode > -1) and self.TranscriptWindow.dlg.editor.get_read_only():
                # Move the Video Start Point to this time position
                self.SetVideoStartPoint(start_timecode)
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
                # Play should always be initiated on Ctrl-P
                self.Play(0)

        # Ctrl-N is Play Next Time-Coded Segment
        elif (c == ord("N")) and event.ControlDown() and loaded:
            # Get the value for the next time code
            start_timecode = self.TranscriptWindow.dlg.editor.NextTimeCode()
            # If there WAS a Next Segment ...
            if (start_timecode > -1) and self.TranscriptWindow.dlg.editor.get_read_only():
                # Move the Video Start Point to this time position
                self.SetVideoStartPoint(start_timecode)
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
                # Play should always be initiated on Ctrl-P
                self.Play(0)

        # Ctrl-S is Stop / Start with Rewind
        elif (c == ord('S')) and event.ControlDown() and loaded:
            if not self.IsPlaying():
                # Explicitly tell Transana to play to the end of the Episode/Clip
                self.SetVideoEndPoint(-1)
            # Play/Pause with auto-rewind
            self.PlayPause(True)

        # Ctrl-T inserts a Time Code
        elif (c == ord('T')) and event.ControlDown() and loaded:
            self.TranscriptWindow.dlg.editor.insert_timecode()

        # Ctrl-. is increases playback speed, if possible
        elif (c == ord('.')) and event.ControlDown() and loaded:
            # Ctrl-period increases playback speed by 10% of normal speed
            self.ChangePlaybackSpeed('faster')

        # Ctrl-, is decreases playback speed, if possible
        elif (c == ord(',')) and event.ControlDown() and loaded:
            # Ctrl-comma decreases playback speed by 10% of normal speed
            self.ChangePlaybackSpeed('slower')

        # Ctrl-M is Shapshot (iMage)
        elif (c == ord('M')) and event.ControlDown() and loaded:
            # If possible ...
            if self.VideoWindow.btnSnapshot:
                # Take a Snapshot
                self.VideoWindow.OnSnapshot(event)

        # Otherwise ...
        else:
            # ... the key press had NOT been processed by this method
            keyProcessed = False

        # Let the calling method know if this method processed the key appropriately
        return keyProcessed

    def InsertTimecodeIntoTranscript(self):
        """ Insert a Timecode into the Transcript(s) """
        # For each Transcript pane in the current Notebook tab ...
        for trWin in self.TranscriptWindow.nb.GetCurrentPage().GetChildren():
            # ... if the transcript is in Edit mode ...
            if not trWin.editor.get_read_only():
                # ... get the transcript's selection ...
                selection = trWin.editor.GetSelection()
                # ... and only if it's a position, not a selection ...
                if selection[0] == selection[1]:
                    # ... then insert the time code.
                    trWin.editor.insert_timecode()

    def InsertSelectionTimecodesIntoTranscript(self, startPos, endPos):
        """ Insert a timed pause into the Transcript """
        self.TranscriptWindow.InsertSelectionTimeCode(startPos, endPos)

    def SetTranscriptEditOptions(self, enable):
        """ Change the Transcript's Edit Mode """
        self.MenuWindow.SetTranscriptEditOptions(enable)

    def ActiveTranscriptReadOnly(self):
        return self.TranscriptWindow.dlg.editor.get_read_only()

    def TranscriptUndo(self, event):
        """ Send an Undo command to the Transcript """
        self.TranscriptWindow.TranscriptUndo(event)

    def TranscriptCut(self, event):
        """ Send a Cut command to the Transcript """
        self.TranscriptWindow.TranscriptCut(event)

    def TranscriptCopy(self, event):
        """ Send a Copy command to the Transcript """
        self.TranscriptWindow.TranscriptCopy(event)

    def TranscriptPaste(self, event):
        """ Send a Paste command to the Transcript """
        self.TranscriptWindow.TranscriptPaste(event)

    def TranscriptCallFormatDialog(self, tabToOpen=0):
        """ Tell the TranscriptWindow to open the Format Dialog """
        self.TranscriptWindow.CallFormatDialog(tabToOpen)

    def TranscriptInsertHyperlink(self, linkType, objNum):
        """ Tell the DocumentWindow to insert a Hyperlink """
        self.TranscriptWindow.InsertHyperlink(linkType, objNum)

    def TranscriptInsertImage(self, fileName = None, snapshotNum = -1):
        """ Tell the TranscriptWindow to insert an image """
        self.TranscriptWindow.InsertImage(fileName, snapshotNum)

    def Help(self, helpContext):
        """ Handles all calls to the Help System """
        # Getting this to work both from within Python and in the stand-alone executable
        # has been a little tricky.  To get it working right, we need the path to the
        # Transana executables, where Help.exe resides, and the file name, which tells us
        # if we're in Python or not.
        (path, fn) = os.path.split(sys.argv[0])
        
        # If the path is not blank, add the path seperator to the end if needed
        if (path != '') and (path[-1] != os.sep):
            path = path + os.sep

        programName = os.path.join(path, 'Help.py')

        if "__WXMAC__" in wx.PlatformInfo:
            # NOTE:  If we just call Help.Help(), you can't actually do the Tutorial because
            # the Help program's menus override Transana's, and there's no way to get them back.
            # instead of the old call:
            
            # Help.Help(helpContext)
            
            # NOTE:  I've tried a bunch of different things on the Mac without success.  It seems that
            #        the Mac doesn't allow command line parameters, and I have not been able to find
            #        a reasonable method for passing the information to the Help application to tell it
            #        what page to load.  What works is to save the string to the hard drive and 
            #        have the Help file read it that way.  If the user leave Help open, it won't get
            #        updated on subsequent calls, but for now that's okay by me.
            
            helpfile = open(os.getenv("HOME") + '/TransanaHelpContext.txt', 'w')
            pickle.dump(helpContext, helpfile)
            helpfile.flush()
            helpfile.close()

            # On OS X 10.4, when Transana is packed with py2app, the Help call stopped working.
            # It seems we have to remove certain environment variables to get it to work properly!
            # Let's investigate environment variables here!
            envirVars = os.environ
            if 'PYTHONHOME' in envirVars.keys():
                del(os.environ['PYTHONHOME'])
            if 'PYTHONPATH' in envirVars.keys():
                del(os.environ['PYTHONPATH'])
            if 'PYTHONEXECUTABLE' in envirVars.keys():
                del(os.environ['PYTHONEXECUTABLE'])

            os.system('open -a TransanaHelp.app')

        else:
            # NOTE:  If we just call Help.Help(), you can't actually do the Tutorial because 
            # modal dialogs prevent you from focussing back on the Help Window to scroll or
            # advance the Tutorial!  Instead of the old call:
        
            # Help.Help(helpContext)

            # we'll use Python's os.spawn() to create a seperate process for the Help system
            # to run in.  That way, we can go back and forth between Transana and Help as
            # independent programs.

            # Make the Help call differently from Python and the stand-alone executable.
            if fn.lower() in ['transana.py', 'unit_test_form_check.py', 'unit_test_search.py']:

                # This works for python 2.7.x
                import subprocess
                subprocess.Popen('"C:\\Program Files (x86)\\Python27-32\\python" "%s" %s' % (programName, helpContext))

                # This worked for Python 2.6.x                
#                # for within Python, we call python, then the Help code and the context
#                os.spawnv(os.P_NOWAIT, 'C:\\Program Files (x86)\\Python27-32\\pythonw', [programName, helpContext])
            else:
                # The Standalone requires a "dummy" parameter here (Help), as sys.argv differs between the two versions.
                os.spawnv(os.P_NOWAIT, path + 'Help', ['Help', helpContext])


    # Private Methods
        
    def LoadVideo(self, currentObj):  # (self, Filename, mediaStart, mediaLength):
        """ This method handles loading a video in the video window and loading the
            corresponding Visualization in the Visualization window. """
        # Get the primary file name
        Filename = currentObj.media_filename
##        # Get the additional files, if any.
##        additionalFiles = currentObj.additional_media_files
##        # If we have a Episode ...
##        if isinstance(currentObj, Episode.Episode):
##            # Initialize the offset value for an Episode to 0, since the 
##            offset = 0
##            # ... the mediaStart is 0 and the mediaLength is the media file length
##            mediaStart = 0
##            mediaLength = currentObj.tape_length
##            # Signal that we have an Episode
##            imgType = 'Episode'
##        # If we have a Clip ...
##        elif isinstance(currentObj, Clip.Clip):
##            # Initialize the offset value for a Clip to the Clip's offset value
##            offset = currentObj.offset
##            # ... the mediaStart is the clip start and the mediaLength is the clip stop - clip start
##            mediaStart = currentObj.clip_start
##            mediaLength = currentObj.clip_stop - currentObj.clip_start
##            # Signal that we have a Clip
##            imgType = 'Clip'
        # See if the primary media file exists
        success = os.path.exists(Filename)
        # If the primary file exists ...
        if success:
            # ... we can iterate through the rest of the media files ...
            for vid in currentObj.additional_media_files:
                # ... to see if they exist as well.
                success = os.path.exists(vid['filename'])
                # As soon as one doesn't exist, we can quit looking.
                if not success:
                    Filename = vid['filename']
                    break
        # If one or more of the Media Files doesn't exist, display an error message.
        if not success:
            # We need a different message for single-user and multi-user Transana if the video file cannot be found.
            if TransanaConstants.singleUserVersion:
                # If it does not exist, display an error message Dialog
                prompt = unicode(_('Media File "%s" cannot be found.\nPlease locate this media file and press the "Update Database" button.\nThen reload the Transcript or Clip that failed.'), 'utf8')
            else:
                # If it does not exist, display an error message Dialog
                prompt = unicode(_('Media File "%s" cannot be found.\nPlease make sure your video root directory is set correctly, and that the video file exists in the correct location.\nThen reload the Transcript or Clip that failed.'), 'utf8')
            dlg = Dialogs.ErrorDialog(self.MenuWindow, prompt % Filename)
            dlg.ShowModal()
            dlg.Destroy()
        else:
            # If the Visualization Window is visible, open the Visualization in the Visualization Window.
            # Loading Visualization first prevents problems with video being locked by Media Player
            # and thus unavailable for wceraudio DLL/Shared Library for audio extraction (in theory).
            # Set the Visualization Window's Visualization Object
            self.VisualizationWindow.SetVisualizationObject(currentObj)

##            # Load the waveform for the appropriate media files with its current start and length.
##            self.VisualizationWindow.load_image(imgType, Filename, additionalFiles, offset, mediaStart, mediaLength)

            # Now that the Visualization is done, load the video in the Video Window
            self.VideoFilename = Filename                # Remember the Video File Name

            # Open the video(s) in the Video Window if the file is found
            self.VideoWindow.open_media_file()
        # Let the calling routine know if we were successful
        return success

    def ClearVisualization(self):
        """ Clear the current selection from the Visualization Window """

        # Clear the currently loaded object, as there is none
        self.currentObj = None
        self.VisualizationWindow.ClearVisualization()

#        self.VisualizationWindow.OnIdle(None)

    def ClearMediaWindow(self):
        """ Clear the media window """
        # Clear the Video Window
        self.VideoWindow.ClearVideo()
        # Clear the Video Filename as well!
        self.VideoFilename = ''
        # Identify the loaded media file
        str = _('Media')
        self.VideoWindow.SetTitle(str)

    def ClearVisualizationSelection(self):
        """ Clear the current selection from the Visualization Window """
        self.VisualizationWindow.ClearVisualizationSelection()

    def ChangeVisualization(self):
        """ Triggers a complete refresh of the Visualization Window.  Needed for changing Visualization Style. """
        # Capture the Transcript Window's cursor position
        self.SaveAllTranscriptCursors()
        # Update the Visualization Window
        self.VisualizationWindow.Refresh()
        # Restore the Transcript Window's cursor
        self.RestoreAllTranscriptCursors()

    def ReplaceVisualizationWindowTextObject(self, tobj):
        """ Change what Text object populates the Visualization window """
        # As long as the Keyword Map is defined ...
        if self.VisualizationWindow.kwMap != None:
            # ... update the text object
            self.VisualizationWindow.kwMap.textObj = tobj

    def UpdateKeywordVisualization(self, textChangeOnly = False):
        """ If the Keyword Visualization is displayed, update it based on something that could change the keywords
            in the display area.  The parameter indicates whether we are just changing text in the object (and therefore
            do not need to update the CONTENTS of the visualization) or whether we need to do the full coding lookup. """
        # Update the Keyword Visualization
        self.VisualizationWindow.UpdateKeywordVisualization(textChangeOnly)

    def UpdateSnapshotWindows(self, mode, event, kw_group, kw_name=None):
        """ mode = 'Update' -- Update all open Snapshot Windows based on a change in a keyword
                   'Detect' -- Return True if the keyword in question is in an open Snapshot in EDIT mode  """
        # We need to update open Snapshots that contain this keyword.
        # Start with a list of all the open Snapshot Windows.
        if mode == 'Update':
            openSnapshotWindows = self.GetOpenSnapshotWindows()
        elif mode == 'Detect':
            openSnapshotWindows = self.GetOpenSnapshotWindows(editableOnly=True)
        else:
            TransanaExceptions.ProgrammingError('Unknown mode "%s" in ControlObjectClass.UpdateSnapshotWindows()' % mode)
        # Initialize a variable that signals whether the keyword to be deleted has been found
        keywordFound = False
        # Interate through the Snapshot Windows
        for win in openSnapshotWindows:
            # Iterate through the Whole Snapshot Keywords
            for kw in win.obj.keyword_list:
                # If we find the Keyword to be deleted ...
                if (kw.keywordGroup == kw_group) and \
                   ((kw_name == None) or (kw.keyword == kw_name)):
                    # ... signal that we found a keyword ...
                    keywordFound = True
                    # ... and stop iterating
                    break
            # Iterate through the Detail Coding for the Snapshot
            for key in win.obj.codingObjects:
                # If we find the Keyword to be deleted ...
                if (win.obj.codingObjects[key]['keywordGroup'] == kw_group) and \
                   ((kw_name == None) or (win.obj.codingObjects[key]['keyword'] == kw_name)):
                    # ... signal that we found a keyword ...
                    keywordFound = True
                    # ... and stop iterating
                    break
            # If we've found the keyword in any Snapshot not in Edit mode ...
            if (mode == 'Update') and (not win.editTool.IsToggled()) and keywordFound:
                # ... update the Snapshot Window
                win.FileClear(event)
                win.FileRestore(event)
                win.OnEnterWindow(event)
            # If we're detecting keywords and we've found one ...
            elif (mode == 'Detect') and (win.editTool.IsToggled()) and keywordFound:
                # ... we don't need to keep looking
                return keywordFound
        # Once we're done with all the Windows, if we're DETECTing ...
        if mode == 'Detect':
            # ... return the results of the detection
            return keywordFound

    def Play(self, setback=False):
        """ This method starts video playback from the current video position. """
        # If we do not already have a cursor position saved, save it
        if self.TranscriptWindow.dlg.editor.cursorPosition == 0:
            self.TranscriptWindow.dlg.editor.SaveCursor()
        # If Setback is requested (Transcription Ctrl-S)
        if setback:
            # Get the current Video position
            videoPos = self.VideoWindow.GetCurrentVideoPosition()
            if type(self.currentObj).__name__ == 'Episode':
                videoStart = 0
            elif type(self.currentObj).__name__ == 'Clip':
                videoStart = self.currentObj.clip_start
            else:
                # Get the current Video marker
                videoStart = self.VideoWindow.GetVideoStartPoint()
            # Assertation: videoPos >= videoStart
            # Find the configured Setback Size (convert to milliseconds)
            setbackSize = TransanaGlobal.configData.transcriptionSetback * 1000
            # If you are further into the video than the Seback Size ...
            if videoPos - videoStart > setbackSize:
                # ... jump back in the video by the setback size
                self.VideoWindow.SetCurrentVideoPosition(videoPos - setbackSize)
            # If the setback would take you to before the beginning of video marker ...
            else:
                # ... jump to the beginning of the video marker
                self.VideoWindow.SetCurrentVideoPosition(videoStart)
        # We need to explicitly set the Clip Endpoint, if it's not known.  (It might be unknown in the ControlObject OR the VideoWindow!)
        # If nothing is loaded, currentObj will be None.  Check to avoid an error.
        if ((self.VideoEndPoint <= 0) or (self.VideoWindow.GetVideoEndPoint() <= 0)) and (self.currentObj != None):
            if type(self.currentObj).__name__ == 'Episode':
                videoEnd = self.VideoWindow.GetMediaLength()
            elif type(self.currentObj).__name__ == 'Clip':
                videoEnd = self.currentObj.clip_stop
            self.SetVideoEndPoint(videoEnd)
        # Play the Video
        self.VideoWindow.Play()

    def Stop(self):
        """ This method stops video playback.  Stop causes the video to be repositioned at the VideoStartPoint. """
        self.VideoWindow.Stop()
        self.RestoreAllTranscriptCursors()

    def Pause(self):
        """ This method pauses video playback.  Pause does not alter the video position, so play will continue from where pause was called. """
        self.VideoWindow.Pause()

    def PlayPause(self, setback=False):
        """ If the video is playing, this pauses it.  If the video is paused, this will make it play. """
        if self.VideoWindow.IsPlaying():
            self.Pause()
        elif self.VideoWindow.IsPaused() or self.VideoWindow.IsStopped():
            self.Play(setback)
        else: # If not playing, paused or stopped, then video not loaded yet
            pass

    def PlayStop(self, setback=False):
        """ If the video is playing, this pauses it.  If the video is paused, this will make it play. """
        if self.VideoWindow.IsPlaying():
            self.Stop()
        elif self.VideoWindow.IsPaused() or self.VideoWindow.IsStopped():
            self.Play(setback)
        else: # If not playing, paused or stopped, then video not loaded yet
            pass

    def PlayLoop(self, startPlay):
        """ Start or stop Looped Playback """
        # Remember if we're supposed to play in a loop
        self.playInLoop = startPlay
        # If we're supposed to play in a loop ...
        if self.playInLoop:
            # Looping doesn't work with Presentation Modes.  You can't stop it without the Visualization Window!
            # Therefore, we need to remember which Presentation Mode we started in.
            if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL):
                self.LoopPresMode = 'All'
            elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO):
                self.LoopPresMode = 'Video'
            elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_TRANS):
                self.LoopPresMode = 'Transcript'
            elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_AUDIO):
                self.LoopPresMode = 'Audio'
            # Now force All Windows Presentation Mode
            self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_ALL, True)
            # ... start playing.
            self.Play()
        # If we're supposed to STOP playing in a loop ...
        else:
            # ... then stop playing!
            self.Stop()
            # Restore the old Presentation Mode
            if self.LoopPresMode == 'Video':
                self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO, True)
            elif self.LoopPresMode == 'Transcript':
                self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_TRANS, True)
            elif self.LoopPresMode == 'Audio':
                self.MenuWindow.menuBar.optionsmenu.Check(MenuSetup.MENU_OPTIONS_PRESENT_AUDIO, True)
            # and clear out the mode that was remembered
            self.LoopPresMode = None

    def IsPlaying(self):
        """ Indicates whether the video is playing or not. """
        return self.VideoWindow.IsPlaying()

    def IsPaused(self):
        """ Indicates whether the video is paused or not. """
        return self.VideoWindow.IsPaused()

    def IsLoading(self):
        """ Indicates whether the video is loading into the Player or not. """
        return self.VideoWindow.IsLoading()

    def GetVideoStartPoint(self):
        """ Return the current Video Starting Point """
        return self.VideoStartPoint
    
    def SetVideoStartPoint(self, TimeCode):
        """ Set the Starting Point for video segment definition.  0 is the start of the video.  TimeCode is the nunber of milliseconds from the beginning. """
        # If we are passed a negative time code ...
        if TimeCode < 0:
            # ... set the time code to 0, the start of the video
            TimeCode = 0
        self.VideoWindow.SetVideoStartPoint(TimeCode)
        self.VideoStartPoint = TimeCode

    def GetVideoEndPoint(self):
        """ Return the current Video Ending Point """
        if self.VideoEndPoint > 0:
            return self.VideoEndPoint
        else:
            return self.VideoWindow.GetMediaLength()

    def SetVideoEndPoint(self, TimeCode):
        """ Set the Stopping Point for video segment definition.  0 is the end of the video.  TimeCode is the nunber of milliseconds from the beginning. """
        self.VideoWindow.SetVideoEndPoint(TimeCode)
        self.VideoEndPoint = TimeCode

    def GetVideoSelection(self):
        """ Return the current video starting and ending points """
        return (self.VideoStartPoint, self.VideoEndPoint)

    def SetVideoSelection(self, StartTimeCode, EndTimeCode, UpdateSelectionText=True):
        """ Set the Starting and Stopping Points for video segment definition.  TimeCodes are in milliseconds from the beginning. """
        # Video selections should only work if we are looking at a Transcript, not a Document
        if self.GetCurrentItemType() == 'Transcript':
            # For each Transcript Window ...
            for trWin in self.TranscriptWindow.nb.GetCurrentPage().GetChildren():
                # ... if the Window is Read Only (not in Edit mode) ...
                if trWin.editor.get_read_only():
                    # Sometime the cursor is positioned at the end of the selection rather than the beginning, which can cause
                    # problems with the highlight.  Let's fix that if needed.
                    if trWin.editor.GetCurrentPos() != trWin.editor.GetSelection()[0]:
                        (start, end) = trWin.editor.GetSelection()
                        trWin.editor.SetCurrentPos(start)
                        trWin.editor.SetAnchor(end)
                    # If Word Tracking is ON ...
                    if TransanaGlobal.configData.wordTracking:
                        # ... highlight the full text of the video selection
                        trWin.editor.scroll_to_time(StartTimeCode)
                        if EndTimeCode > 0:
                            trWin.editor.select_find(str(EndTimeCode))

            if EndTimeCode <= 0:
                if type(self.currentObj).__name__ == 'Episode':
                    EndTimeCode = self.VideoWindow.GetMediaLength()
                elif type(self.currentObj).__name__ == 'Clip':
                    EndTimeCode = self.currentObj.clip_stop
            self.SetVideoStartPoint(StartTimeCode)
            self.SetVideoEndPoint(EndTimeCode)
            # The SelectedEpisodeClips window was not updating on the Mac.  Therefore, this was added,
            # even if it might be redundant on Windows.
            if (not self.IsPlaying()) or (self.TranscriptWindow.UpdatePosition(StartTimeCode)):
                if self.DataWindow.SelectedDataItemsTab != None:
                    self.DataWindow.SelectedDataItemsTab.Refresh(StartTimeCode)
            # If we should update Selection Text (true at all times other than when clearing the Visualization Window ...)
            if UpdateSelectionText:
                # Update the Selection Text in the current Transcript Window.  But it needs just a tick before the cursor position is set correctly.
                wx.CallLater(50, self.UpdateSelectionTextLater)

    def UpdatePlayState(self, playState):
        """ When the Video Player's Play State Changes, we may need to adjust the Screen Layout
            depending on the Presentation Mode settings. """

##        print "ControlObjectClass.UpdatePlayState():", playState, TransanaConstants.MEDIA_PLAYSTATE_NONE, \
##          TransanaConstants.MEDIA_PLAYSTATE_STOP, TransanaConstants.MEDIA_PLAYSTATE_PAUSE, \
##          TransanaConstants.MEDIA_PLAYSTATE_PLAY

        # If media playback is stopping and we're supposed to be playing in a loop ...
        if (playState == TransanaConstants.MEDIA_PLAYSTATE_STOP) and self.playInLoop:
            # ... then re-start media playback
            self.Play()
        # If the video is STOPPED, return all windows to normal Transana layout
        elif (playState in [TransanaConstants.MEDIA_PLAYSTATE_STOP, TransanaConstants.MEDIA_PLAYSTATE_PAUSE]):
            # If we don't have a PlayAllClips Window or the PlayAllClips Shutdown has been triggered ...
            if (self.PlayAllClipsWindow == None) or self.shutdownPlayAllClips:
                # When Play is intiated (below), the positions of windows gets saved if they are altered by Presentation Mode.
                # If this has happened, we need to put the screen back to how it was before when Play is stopped.
                if len(self.WindowPositions) != 0:
                    # Reset the AutoArrange (which was temporarily disabled for Presentation Mode) variable based on the Menu Setting
                    TransanaGlobal.configData.autoArrange = self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_AUTOARRANGE)
                    # Reposition the Video Window to its original Position (self.WindowsPositions[2])
                    self.VideoWindow.SetDims(self.WindowPositions[2][0], self.WindowPositions[2][1], self.WindowPositions[2][2], self.WindowPositions[2][3])
                    # Show the Menu Bar
                    self.MenuWindow.Show(True)
                    # Show the Visualization Window
                    self.VisualizationWindow.Show(True)
                    # Show the Transcript Window
                    self.TranscriptWindow.Show(True)
                    # Show the Video Window
                    self.VideoWindow.Show(True)
                    # Show the Data Window
                    self.DataWindow.Show(True)
                    # Clear the saved Window Positions, so that if they are moved, the new settings will be saved when the time comes
                    self.WindowPositions = []
                # Reset the Transcript Cursors
                self.RestoreAllTranscriptCursors()
                # Reset the Shutdown flag
                self.shutdownPlayAllClips = False
            # if we have a PlayAllClipsWindow but its shutdown has not been triggered ...
            else:
                # ... see if we're playing the LAST clip now.
                if (('wxMSW' in wx.PlatformInfo) and \
                    (self.PlayAllClipsWindow.clipNowPlaying == len(self.PlayAllClipsWindow.clipList))) or \
                   (('wxMac' in wx.PlatformInfo) and \
                    (self.PlayAllClipsWindow.clipNowPlaying == len(self.PlayAllClipsWindow.clipList) - 1) and \
                    (self.PlayAllClipsWindow.btnPlayPause.GetLabel() == _("Pause"))):
                    # If so, set the Play All Clips Shutdown flag to True!
                    self.shutdownPlayAllClips = True

        # If the video is PLAYED, adjust windows to the desired screen layout,
        # as indicated by the Presentation Mode selection
        elif playState == TransanaConstants.MEDIA_PLAYSTATE_PLAY:
            # If we are starting up from the Video Window, save the Transcript Cursor.
            # Detecting that the Video Window has focus is hard, as there are different video window implementations on
            # different platforms.  Therefore, let's see if it's NOT the Transcript or the Waveform, which are easier to
            # detect.
            if (type(self.MenuWindow.FindFocus()) != type(self.TranscriptWindow.dlg.editor)) and \
               ((self.MenuWindow.FindFocus()) != (self.VisualizationWindow.waveform)):
                self.TranscriptWindow.dlg.editor.SaveCursor()
            # See if Presentation Mode is NOT set to "All Windows" and do all changes common to the other Presentation Modes
            if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_ALL) == False:
                # See if we have already noted the Window Positions.
                if len(self.WindowPositions) == 0:
                    # If not...
                    # Temporarily disable AutoArrange, as it interferes with Presentation Mode
                    TransanaGlobal.configData.autoArrange = False
                    # Save the Window Positions prior to Presentation Mode rearrangement
                    self.WindowPositions = [self.MenuWindow.GetRect(),
                                            self.VisualizationWindow.GetDimensions(),
                                            self.VideoWindow.GetDimensions(),
                                            self.TranscriptWindow.GetDimensions(),
                                            self.DataWindow.GetDimensions()]
                # If we're on Windows ...
                if 'wxMSW' in wx.PlatformInfo:
                    # ... hide the Menu Window
                    self.MenuWindow.Show(False)
                # Hide the Visualization Window
                self.VisualizationWindow.Show(False)
                # Hide the Data Window
                self.DataWindow.Show(False)
                # Determine which monitor to use and get its size and position
                if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
                    primaryScreen = TransanaGlobal.configData.primaryScreen
                else:
                    primaryScreen = 0
                (left, top, width, height) = wx.Display(primaryScreen).GetClientArea()
                # See if Presentation Mode is set to "Video Only"
                if self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_VIDEO):
                    # Hide the Transcript Windows
                    self.TranscriptWindow.Show(False)
                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Video Window to take up almost the whole Client Display area
                        self.VideoWindow.SetDims(left + 1, top + 1, width - 2, height - 61)
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 1
                        self.PlayAllClipsWindow.yPos = height - 58
                        # We need a bit more adjustment on the Mac
                        if ('wxMac' in wx.PlatformInfo) and (TransanaGlobal.configData.primaryScreen == 0):
                            self.PlayAllClipsWindow.yPos += 24
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 2, 56))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()

                    # If there's NO play all clips window within Video Only presentation mode ...
                    else:
                        # Set the Video Window to take up almost the whole Client Display area
                        self.VideoWindow.SetDims(left + 1, top + 1, width - 2, height - 2)

                # See if Presentation Mode is set to "Video and Transcript"
                elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_TRANS):
                    # Set the screen proportions, currently 70% video, 30% transcript
                    dividePt = 70 / 100.0
                    # We need to make a slight adjustment for the Mac for the menu height
                    if 'wxMac' in wx.PlatformInfo:
                        height += TransanaGlobal.menuHeight
                    # Set the Video Window to take up the top portion of the Client Display Area
                    self.VideoWindow.SetDims(left + 1, top + 2, width - 2, int(dividePt * height) - 3)
                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 1
                        self.PlayAllClipsWindow.yPos = int(dividePt * height) + 1
                        if 'wxMac' in wx.PlatformInfo:
                            self.PlayAllClipsWindow.yPos -= 20
                        # We need a bit more adjustment on Windows for Video and Transcript mode
                        elif ('wxMSW' in wx.PlatformInfo):
                            self.PlayAllClipsWindow.yPos -= 28
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 2, 56))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()
                    # Set the Transcript Window to take up the bottom portion of the Client Display Area
                    self.TranscriptWindow.SetDims(left + 1, int(dividePt * height) + 1, width - 2, int((1.0 - dividePt) * height) - 2)
                    # if we're NOT using the Rich Text Ctrl (ie. we are using the Styled Text Ctrl) ...
                    if not TransanaConstants.USESRTC:
                        # Set the Transcript Zoom Factor
                        if 'wxMac' in wx.PlatformInfo:
                            zoomFactor = 20
                        else:
                            zoomFactor = 14
                        # Zoom in the Transcript window to make the text larger
                        self.TranscriptWindow.dlg.editor.SetZoom(zoomFactor)

                # See if Presentation Mode is set to "Audio and Transcript"
                elif self.MenuWindow.menuBar.optionsmenu.IsChecked(MenuSetup.MENU_OPTIONS_PRESENT_AUDIO):
                    # Hide the Video Window to get Audio with no Video
                    self.VideoWindow.Show(False)
                    # We need to make a slight adjustment for the Mac for the menu height
                    if 'wxMac' in wx.PlatformInfo:
                        height += TransanaGlobal.menuHeight
                    # Set the Transcript Window to take up the entire Client Display Area
                    self.TranscriptWindow.SetDims(left + 1, top + 1, width - 2, height - 2)

                    # If there is a PlayAllClipsWindow, reset it's size and layout
                    if self.PlayAllClipsWindow != None:
                        # Set the Window Position in the PlayAllClips Dialog
                        self.PlayAllClipsWindow.xPos = left + 1
                        self.PlayAllClipsWindow.yPos = top
                        if 'wxMac' in wx.PlatformInfo:
                            winHeight = self.TranscriptWindow.GetSizeTuple()[1] - self.TranscriptWindow.GetClientSizeTuple()[1] - 10
                        else:
                            ## Removed because we were getting too large a window, especially in Arabic.
                            ## winHeight = self.TranscriptWindow.GetSizeTuple()[1] - self.TranscriptWindow.GetClientSizeTuple()[1] + 30

                            winHeight = 56
                            
                        self.PlayAllClipsWindow.SetRect(wx.Rect(self.PlayAllClipsWindow.xPos, self.PlayAllClipsWindow.yPos, width - 2, winHeight))
                        # Make the PlayAllClipsWindow the focus
                        self.PlayAllClipsWindow.SetFocus()

    def ChangePlaybackSpeed(self, direction):
        """ Change the media playback speed on the fly """
        # Pass the request through to the Video Window
        self.VideoWindow.ChangePlaybackSpeed(direction)

    def GetDatabaseDims(self):
        """ Return the dimensions of the Database control. Note that this only returns the Database Tree Tab location.  """
        # Determine the Screen Position of the top left corner of the Tree Control
        (treeLeft, treeTop) = self.DataWindow.DBTab.tree.ClientToScreenXY(1, 1)
        # Determine the width and height of the tree control
        (width, height) = self.DataWindow.DBTab.tree.GetSizeTuple()
        # Return the Database Tree Tab position and size information
        return (treeLeft, treeTop, width, height)

    def GetTranscriptDims(self):
        """ Return the dimensions of the transcript control.  Note that this only includes the transcript itself
        and not the whole Transcript window (including toolbars, etc). """
        return self.TranscriptWindow.GetTranscriptDims()

    def GetCurrentTranscriptObject(self):
        """ Returns a Transcript Object for the Transcript currently loaded in the Transcript Editor """
        return self.TranscriptWindow.GetCurrentTranscriptObject()

    def GetTranscriptSelectionInfo(self):
        """ Returns information about the current selection in the transcript editor """
        # We need to know the time codes that bound the current selection
        (startTime, endTime) = self.TranscriptWindow.dlg.editor.get_selected_time_range()
        # we need to know the text of the current selection
        # If it's blank, we need to send a blank rather than RTF for nothing
        (startPos, endPos) = self.TranscriptWindow.dlg.editor.GetSelection()
        # If there's no current selection ...
        if startPos == endPos:
            # ... get the text between the nearest time codes.
            (st, end, text) = self.TranscriptWindow.dlg.editor.GetTextBetweenTimeCodes(startTime, endTime)
        else:
            if TransanaConstants.USESRTC:
                text = self.TranscriptWindow.dlg.editor.GetFormattedSelection('XML', selectionOnly=True)
            else:
                text = self.TranscriptWindow.dlg.editor.GetRTFBuffer(select_only=1)
        # We also need to know the number of the original Transcript Record
        if self.TranscriptWindow.dlg.editor.TranscriptObj.clip_num == 0:
            # If we have an Episode Transcript, we need the Transcript Number
            originalTranscriptNum = self.TranscriptWindow.dlg.editor.TranscriptObj.number
        else:
            # If we have a Clip Transcript, we need the original Transcript Number, not the Clip Transcript Number.
            # We can get that from the ControlObject's "currentObj", which in this case will be the Clip!
            originalTranscriptNum = self.currentObj.transcripts[self.activeTranscript].source_transcript
        return (originalTranscriptNum, startTime, endTime, text)

    def AddQuoteToOpenDocument(self, tmpQuote):
        """ When a Quote is added, we need to make sure the Document, if open, is informed about it!! """
        # For each Page (Tab) in the Transcript Window ...
        for page in range(self.TranscriptWindow.nb.GetPageCount()):
            # For each Splitter Pane on the Page ...
            for pane in self.TranscriptWindow.nb.GetPage(page).GetChildren():
                # Do we have a Document, and is it the correct Document ...
                if isinstance(pane.editor.TranscriptObj, Document.Document) and (pane.editor.TranscriptObj.number == tmpQuote.source_document_num):
                    # Add the record directly to the open Document Object.  It's already been added to the database!
                    pane.editor.TranscriptObj.quote_dict[tmpQuote.number] = (tmpQuote.start_char, tmpQuote.end_char)
                    # Once found, we can stop looking
                    break

    def SetCurrentDocumentPosition(self, insertPos, sel, fromEditor=False):
        """ Set the Cursor Position in the current Document in the Transcript Window """

##        print
##        print "ControlObjectClass.SetCurrentDocumentPosition(%s, %s)" % (insertPos, fromEditor)
##        print
        
        if sel == (-2, -2) or (sel[1] - sel[0] == 1):
            if not fromEditor:
                # We MUST use insertPos + 1.  A position doesn't show in the transcript, while a Selection does!
                self.SetCurrentDocumentSelection(insertPos, insertPos + 1, fromEditor)
            else:
                self.SetCurrentDocumentSelection(insertPos, insertPos, fromEditor)
        else:
            self.SetCurrentDocumentSelection(sel[0], sel[1], fromEditor)
        # Place the Visualization Cursor
        self.VisualizationWindow.startPoint = insertPos
        self.VisualizationWindow.endPoint = insertPos
        self.VisualizationWindow.redrawWhenIdle = True
        # If the Selected items tab exists and is showing ...
        if (self.DataWindow.SelectedDataItemsTab != None) and self.DataWindow.SelectedDataItemsTab.IsShown():
            # ... update it for the current position
            self.DataWindow.SelectedDataItemsTab.Refresh(insertPos, (-2, -2))

    def SetCurrentDocumentSelection(self, startPos, endPos, fromEditor=False):
        """ Set the Selection in the current Document in the Transcript Window """

##        print
##        print "ControlObjectClass.SetCurrentDocumentSelection(%s, %s, %s)" % (startPos, endPos, fromEditor)
##        print

        # If our Document is Read Only ...  (We should not change position or selection in Edit mode.)
        if self.TranscriptWindow.dlg.editor.get_read_only():
            # If this call did NOT come from the Editor control ...
            if not fromEditor:
                # Set the Selection in the Document Window
                self.TranscriptWindow.dlg.editor.SetSelection(startPos, endPos)
                # Make sure the current selection is visible on screen
                self.TranscriptWindow.dlg.editor.ShowCurrentSelection()
        # Update the Visualization Window
        self.VisualizationWindow.SetDocumentSelection(startPos, endPos)
        # If the Selected items tab exists and is showing ...
        if (self.DataWindow.SelectedDataItemsTab != None) and self.DataWindow.SelectedDataItemsTab.IsShown():
            # ... update it for the current position
            self.DataWindow.SelectedDataItemsTab.Refresh(startPos, (startPos, endPos))

    def DeleteVisualizationInfo(self, objKey):
        """ Pass-through used to delete Visualization Information when closing a Document Notebook Page """

        try:
            # If our key is for a Transcript ...
            if objKey[0] == Transcript.Transcript:
                # ... load the Transcript object here
                tmpObj = Transcript.Transcript(objKey[1])
                # If we have an Episode Transcript ...
                if tmpObj.clip_num == 0:
                    # Update the Object Key to be the Episode or Clip the transcript is connected to
                    objKey = (Episode.Episode, tmpObj.episode_num)
                # If we have a Clip Transcript ...
                else:
                    # Update the Object Key to be the Episode or Clip the transcript is connected to
                    objKey = (Clip.Clip, tmpObj.clip_num)
            # Pass the Delete request on to the Visualization Window
            self.VisualizationWindow.DeleteVisualizationInfo(objKey)
        except TransanaExceptions.RecordNotFoundError, e:

            # If you delete a CLIP on ANOTHER computer that is also open (but not in Edit mode) on THIS computer,
            # then Transana won't be able to find the Transcript Record.  In that case, we have to delete the
            # VizualizationWindow's VizualizationInfo manually!

            for key in self.VisualizationWindow.VisualizationInfo.keys():
                if isinstance(key[0], Transcript.Transcript):
                    del(self.VizualizationWindow.VizualizationInfo[key])
                    break

    def HighlightQuoteInCurrentDocument(self, quote):
        """ Highlight the indicated Quote in the current document
            (This compensates for changed Quote positions due to document editing! """
        try:
            # Get the current (possibly changed) quote start and end positions for the desired quote
            (startPos, endPos) = self.TranscriptWindow.dlg.editor.TranscriptObj.quote_dict[quote.number]
            # Set the Document Selection to match the quote
            self.SetCurrentDocumentSelection(startPos, endPos)
        except:

            print "ControlObjectClass.HighlightQuoteInCurrentDocument():  Cannot find Quote positions."
            print sys.exc_info()[0]
            print sys.exc_info()[1]

    def GetDocumentPosition(self):
        """ Returns Position and Selection information for the currently selected Document Object """
        # If we have a document or a quote ...
        if isinstance(self.currentObj, Document.Document) or isinstance(self.currentObj, Quote.Quote):
            # return the current position and the selection information
            return (self.TranscriptWindow.dlg.editor.GetCurrentPos(), self.TranscriptWindow.dlg.editor.GetSelection())
        # If we don't have a document ...
        else:
            # ... we don't have data to return
            return (-1, (-2, -2))
        
    def GetDocumentSelectionInfo(self):
        """ Returns information about the current selection in the document editor """
        # We need a list of the whitespace character's ordinal values
        whitespace = []
        for ch in string.whitespace:
            whitespace.append(ord(ch))
        # We also need the "period" to be part of that list
        whitespace.append(ord('.'))
        # We need to know the text of the current selection
        # If it's blank, we need to send a blank rather than RTF for nothing
        (startChar, endChar) = self.TranscriptWindow.dlg.editor.GetSelection()
        # If there's no current selection ...
        if startChar == endChar:
            # ... get the current cursor position
            currentPos = self.TranscriptWindow.dlg.editor.GetCurrentPos()
            # Look for the sentence START
            startChar = currentPos
            while not (self.TranscriptWindow.dlg.editor.GetCharAt(startChar) in [ord('.'), ord('\n')]) and \
                  (startChar > 0) and \
                  (currentPos - startChar < 250):
                startChar -= 1
            # Remove the preceeding period and whitespace from the start of the selection.
            while self.TranscriptWindow.dlg.editor.GetCharAt(startChar) in whitespace:
                startChar += 1

            # Look for the sentence END
            endChar = currentPos
            while not (self.TranscriptWindow.dlg.editor.GetCharAt(endChar) in [ord('.'), ord('\n')]) and \
                  (endChar < self.TranscriptWindow.dlg.editor.GetLastPosition()) and \
                  (endChar - currentPos < 250):
                endChar += 1
            # Be sure to include the period in the selection
            if self.TranscriptWindow.dlg.editor.GetCharAt(endChar) == ord('.'):
                endChar += 1

            # Set the Transcript/Document selection to the sentence
            self.TranscriptWindow.dlg.editor.SetSelection(startChar, endChar)
            # Get the selected text in XML format
            text = self.TranscriptWindow.dlg.editor.GetFormattedSelection('XML', selectionOnly=True)
            # Restore the original cursor position
            self.TranscriptWindow.dlg.editor.SetCurrentPos(currentPos)
        # If there is a selection ...
        else:
            # ... get the selected text in XML format
            text = self.TranscriptWindow.dlg.editor.GetFormattedSelection('XML', selectionOnly=True)
        # Initialize originalDocumentNum
        originalDocumentNum = 0
        # We also need to know the number of the original Document Record
        if isinstance(self.TranscriptWindow.GetCurrentObject(), Document.Document):
            # If we have a Document, we need the Document Number
            originalDocumentNum = self.TranscriptWindow.GetCurrentObject().number
        elif isinstance(self.TranscriptWindow.GetCurrentObject(), Quote.Quote):
            # If we have a Quote, we need the Source Document Number, not the Quote Number.
            originalDocumentNum = self.TranscriptWindow.GetCurrentObject().source_document_num
            # Adjust the Start and End Chars for the start point of the Quote.  (This is an approximation, and could be incorrect,
            # but it's the best we can do!)
            startChar += self.TranscriptWindow.GetCurrentObject().start_char
            endChar += self.TranscriptWindow.GetCurrentObject().start_char
        return (originalDocumentNum, startChar, endChar, text)

    def GetMultipleTranscriptSelectionInfo(self):
        """ Returns information about the current selection(s) in the transcript editor(s) """
        # Initialize a list for the function results
        results = []
        # Iterate through the transcript windows
        for trWindow in self.TranscriptWindow.nb.GetCurrentPage().GetChildren():
            # We need to know the time codes that bound the current selection in the current transcript window
            (startTime, endTime) = trWindow.editor.get_selected_time_range()
            # If start is 0 ...
            if startTime == 0:
                # ... and we're in a Clip ...
                if isinstance(self.currentObj, Clip.Clip):
                    # ... then the start should be the Clip Start
                    startTime = self.currentObj.clip_start
            # If there is not following time code ...
            if endTime <= 0:
                # ... and we're in a Clip ...
                if isinstance(self.currentObj, Clip.Clip):
                    # ... use the Clip Stop value ...
                    endTime = self.currentObj.clip_stop
                # ... otherwise ...
                else:
                    # ... use the length of the media file
                    endTime = self.GetMediaLength(entire = True)
            # we need to know the text of the current selection in the current transcript window
            # If it's blank, we need to send a blank rather than RTF for nothing
            (startPos, endPos) = trWindow.editor.GetSelection()
            if startPos == endPos:
                text = ''
            else:
                #text = trWindow.editor.GetRTFBuffer(select_only=1)
                if TransanaConstants.USESRTC:
                    text = trWindow.editor.GetFormattedSelection('XML', selectionOnly=True)
                else:
                    text = trWindow.editor.GetRTFBuffer(select_only=1)
            # We also need to know the number of the original Transcript Record.  If we have an Episode ....
            if trWindow.editor.TranscriptObj.clip_num == 0:
                # ... we need the Transcript Number, which we can get from the Transcript Window's editor's Transcript Object
                originalTranscriptNum = trWindow.editor.TranscriptObj.number
            # If we have a Clip ...
            else:
                # ... we need the original Transcript Number, not the Clip Transcript Number.
                # We can get that from the ControlObject's "currentObj", which in this case will be the Clip!
                # We have to pull the source_transcript value from the correct transcript number!
                originalTranscriptNum = trWindow.editor.TranscriptObj.source_transcript
            # Now we can place this transcript's results into the Results list
            results.append((originalTranscriptNum, startTime, endTime, text))
        return results

    def GetDatabaseTreeTabObjectNodeType(self):
        """ Get the Node Type of the currently selected object in the Database Tree in the Data Window """
        return self.DataWindow.DBTab.tree.GetObjectNodeType()

    def SetDatabaseTreeTabCursor(self, cursor):
        """ Change the shape of the cursor for the database tree in the data window """
        self.DataWindow.DBTab.tree.SetCursor(wx.StockCursor(cursor))

    def GetVideoPosition(self):
        """ Returns the current Time Code from the Video Window """
        return self.VideoWindow.GetCurrentVideoPosition()

    def GetVideoCheckboxDataForClips(self, videoPos):
        """ Return the data about the media players checkboxes needed for Clip Creation """
        return self.VideoWindow.GetVideoCheckboxDataForClips(videoPos)

    def VideoCheckboxChange(self):
        """ Detect and adjust to changes in the status of the media player checkboxes """
        # If the check boxes change, the visualization should reflect that.  To accomplish that, we just need to
        # have the Visualization redraw itself when there is idle time.
        self.VisualizationWindow.redrawWhenIdle = True
        
    def UpdateVideoPosition(self, currentPosition):
        """ This method accepts the currentPosition from the video window and propagates that position to other objects """
        # There's a weird glitch with Play All Clips when switching from one multi-transcript clip to the next.
        # Somehow, activeTranscript is getting set to a TranscriptWindow that hasn't yet been created, and that's
        # causing a problem HERE.  The following lines of code fix it.  I haven't been able to track down the cause.
##        if self.activeTranscript >= len(self.TranscriptWindow):
##            # We need to reset the activeTranscript to 0.  It gets reset later when this line is triggered.
##            self.activeTranscript = 0

        # If we do not already have a cursor position saved, and there is a defined cursor position, save it
        if (self.TranscriptWindow.dlg.editor.cursorPosition == 0) and \
           (self.TranscriptWindow.dlg.editor.GetCurrentPos() != 0) and \
           (self.TranscriptWindow.dlg.editor.GetSelection() != (0, 0)):
            self.TranscriptWindow.dlg.editor.SaveCursor()
            
        if self.VideoEndPoint > 0:
            mediaLength = self.VideoEndPoint - self.VideoStartPoint
        else:
            mediaLength = self.VideoWindow.GetMediaLength()
        self.VisualizationWindow.UpdatePosition(currentPosition)

        # Update Transcript position.  If Transcript position changes,
        # then also update the selected Clips tab in the Data window.
        # NOTE:  self.IsPlaying() check added because the SelectedEpisodeClips Tab wasn't updating properly
        if (not self.IsPlaying()) or (self.TranscriptWindow.UpdatePosition(currentPosition)):
            if self.DataWindow.SelectedDataItemsTab != None:
                self.DataWindow.SelectedDataItemsTab.Refresh(currentPosition)

        self.TranscriptWindow.UpdatePosition(currentPosition)
        # Update the Selection Text
        self.UpdateSelectionTextLater()

    def UpdateSelectionTextLater(self, panelNum=-1):
        """ Update the Selection Text after the application has had a chance to update the Selection information """

##        print "ControlObjectClass.UpdateSelectionTextLater(): updated  ???????????????????????"
        
        if panelNum > -1:
            originalPanel = self.activeTranscript
            self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(panelNum)

        # Get the current selection
        selection = self.TranscriptWindow.dlg.editor.GetSelection()
        # If we have a point rather than a selection ...
        if selection[0] == selection[1]:
            # We don't need a label
            lbl = ""
        # If we have a selection rather than a point ...
        else:
            # ... we first need to get the time range of the current selection.
            (start, end) = self.TranscriptWindow.dlg.editor.get_selected_time_range()
            # If start is 0 ...
            if start == 0:
                # ... and we're in a Clip ...
                if isinstance(self.currentObj, Clip.Clip):
                    # ... then the start should be the Clip Start
                    start = self.currentObj.clip_start
            # If there is not following time code ...
            if end <= 0:
                # ... and we're in a Clip ...
                if isinstance(self.currentObj, Clip.Clip):
                    # ... use the Clip Stop value ...
                    end = self.currentObj.clip_stop
                # ... otherwise ...
                else:
                    # ... use the length of the media file
                    end = self.GetMediaLength(entire = True)
            # Then we build the label.
            lbl = unicode(_("Selection:  %s - %s"), 'utf8')
            lbl = lbl % (Misc.time_in_ms_to_str(start), Misc.time_in_ms_to_str(end))
        # Now display the label on the Transcript Window.  But only if it's changed (for optimization during HD playback)
        if self.TranscriptWindow.selectionText.GetLabel() != lbl:
            self.TranscriptWindow.UpdateSelectionText(lbl)

        if panelNum > -1:
            self.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(originalPanel)

    def GetMediaLength(self, entire = False):
        """ This method returns the length of the entire video/media segment """
        try:
            if not(entire): # Return segment length
                # if the end point is not defined (possibly due to media length not yet being available) 
                if self.VideoEndPoint <= 0:
                    # Get the length of the longest adjusted media file
                    videoLength = self.VideoWindow.GetMediaLength()
                    # Subtract the video start point, to get segment length
                    mediaLength = videoLength - self.VideoStartPoint

                # If the video end point is defined ...
                else:
                    # ... as long as the length is positive ...
                    if self.VideoEndPoint - self.VideoStartPoint > 0:
                        # ... get the current segment length
                        mediaLength = self.VideoEndPoint - self.VideoStartPoint
                    # If the length is negative ...
                    else:
                        # ... use the total media length as the end point
                        mediaLength = self.VideoWindow.GetMediaLength() - self.VideoStartPoint

                # Sometimes video files don't know their own length because it hasn't been available before.
                # This seems to be a good place to detect and correct that problem before it starts to cause problems,
                # such as in the Keyword Map.

                # First, let's see if an episode is currently loaded that doesn't have a proper length.
                if (isinstance(self.currentObj, Episode.Episode) and \
                   (self.currentObj.media_filename == self.VideoFilename) and \
                   (self.currentObj.tape_length <= 0) and \
                    (self.VideoWindow.mediaPlayers[0].GetMediaLength() > 0)):
                    # Start exception handling, so record lock errors can be ignored
                    try:
                        # Try to lock the record
                        self.currentObj.lock_record()
                        # Get the media length from the first Video Window, not the VideoWindow object, which reports longest adjusted media file length.
                        self.currentObj.tape_length = self.VideoWindow.mediaPlayers[0].GetMediaLength()
                        # for each additional media window ...
                        for x in range(1, len(self.VideoWindow.mediaPlayers)):
                            # ... get the length for each additional media file.  (We need them all.)
                            self.currentObj.additional_media_files[x - 1]['length'] = self.VideoWindow.mediaPlayers[x].GetMediaLength()
                        # Save the object
                        self.currentObj.db_save()
                        # Unlock the record
                        self.currentObj.unlock_record()
                    # If an exception occurs (most likely a Record Lock exception)
                    except:
                        # it can be ignored.
                        pass

                # Return the calculated value
                return mediaLength
            # If the entire length was requested ...
            else:
                # Return length of longest adjusted media file
                return self.VideoWindow.GetMediaLength()
        except:
            # If an exception is raised, most likely we're shutting down and have lost the VideoWindow.  Just return 0.
            return 0

    def GetDocumentLength(self):
        """ Return the length in characters of the currently-loaded Document / Quote object """
        # Get the current length of the active document's editor control.  (This allows for changes in length due
        # to editing.)
        return self.TranscriptWindow.dlg.editor.GetLength()
        
    def UpdateVideoWindowPosition(self, left, top, width, height):
        """ This method receives screen position and size information from the Video Window and adjusts all other windows accordingly """
        if TransanaGlobal.configData.autoArrange:
            # Determine the values for potential screen position adjustment
            (adjustX, adjustY, adjustW, adjustH) = wx.Display(TransanaGlobal.configData.primaryScreen).GetClientArea()
            # This method behaves differently when there are multiple Video Players open
            if (self.currentObj != None)  and \
               (isinstance(self.currentObj, Episode.Episode) or (isinstance(self.currentObj, Clip.Clip))) and \
               (len(self.currentObj.additional_media_files) > 0):
                # Get the current dimensionf of the Video Window
                (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                # Update other windows based on this information
                self.UpdateWindowPositions('Video', wleft, wtop + wheight)
            else:

                # Get the current dimensions of the Video Window
                (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                # Update other windows based on this information

                if DEBUG:
                    print
                    print "Call 1", 'Video', wleft-1, wtop + wheight
                    print 'Video:', self.VideoWindow.GetDimensions()
                    print 'Data: ', self.DataWindow.GetDimensions()
                    print
                
                self.UpdateWindowPositions('Video', wleft - 1, YLower=wtop + wheight)

            # Play All Clips Window matches the Data Window's WIDTH
            if self.PlayAllClipsWindow != None:
                (parentLeft, parentTop, parentWidth, parentHeight) = self.DataWindow.GetRect()
                (left, top, width, height) = self.PlayAllClipsWindow.GetRect()
                if (parentWidth != width) or (parentLeft != left):
                    self.PlayAllClipsWindow.SetDimensions(parentLeft, top, parentWidth, height)

    def UpdateWindowPositions(self, sender, X, YUpper=-1, YLower=-1):
        """ This method updates all window sizes/positions based on the intersection point passed in.
            X is the horizontal point at which the visualization and transcript windows end and the
            video and data windows begin.
            YUpper is the vertical point where the visualization window ends and the transcript window begins.
            YLower is the vertical point where the video window ends and the data window begins. """

        # NOTE:  This routine was originally written when only one media file could be displayed.  That's why it's
        #        a little more awkward in the multiple-video case.  The data passed in assumes a single video, and
        #        we have to adjust for that.

        if DEBUG:
            print
            print "ControlObjectClass.UpdateWindowPositions(1):", sender, X, YUpper, YLower, TransanaGlobal.resizingAll
            print 'Visualization:', self.VisualizationWindow.GetDimensions()
            print 'Transcript:   ', self.TranscriptWindow.GetDimensions()
            print 'Video:        ', self.VideoWindow.GetDimensions()
            print 'Data:         ', self.DataWindow.GetDimensions()
            print

        # If Transana is in the process of shutting down ...
        if self.shuttingDown:
            # ... we can skip this method
            return

        # If there are more than one monitors on this system and the system is configured to use one of them ...
        if TransanaGlobal.configData.primaryScreen < wx.Display.GetCount():
            # ... use the configured monitor as the primary screen
            primaryScreen = TransanaGlobal.configData.primaryScreen
        # If not ...
        else:
            # ... use the default primary monitor
            primaryScreen = 0
        # Determine the values for potential screen position adjustment
        (adjustX, adjustY, adjustW, adjustH) = wx.Display(primaryScreen).GetClientArea()

        # Fix for Linux
        if 'wxGTK' in wx.PlatformInfo:
            adjustW = wx.Display(primaryScreen).GetClientArea()[2]   # GetGeometry()[2]

        if DEBUG:
            print "Adjusts:", (adjustX, adjustY, adjustW, adjustH)

        # We need to adjust the Window Positions to accomodate multiple transcripts!
        # Basically, if we are not in the "first" transcript, we need to substitute the first transcript's
        # "Top position" value for the one sent by the active window.
        if (sender == 'Transcript'):
            YUpper = self.TranscriptWindow.GetRect()[1] - 1
        # If Auto-Arrange is enabled, resizing one window may alter the positioning of others.
        if TransanaGlobal.configData.autoArrange:
            # If YUpper is NOT passed in ...
            if YUpper == -1:
                # Set it to the BOTTOM of the Visualization Window
                (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
                YUpper = wheight + wtop
            # if YLower is NOT passed in ...
            if YLower == -1:
                # Set it to the BOTTOM of the Video Window
                (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                YLower = wheight + wtop
                
            # This method behaves differently when there are multiple Video Players open
            if (self.currentObj != None) and \
               (isinstance(self.currentObj, Episode.Episode) or (isinstance(self.currentObj, Clip.Clip))) and \
               (len(self.currentObj.additional_media_files) > 0):
                # Get the current dimensions of all windows
                visualDims = self.VisualizationWindow.GetDimensions()
                transcriptDims = self.TranscriptWindow.GetDimensions()
                videoDims = self.VideoWindow.GetDimensions()
                dataDims = self.DataWindow.GetDimensions()
                # If the Visualization window has been changed ...
                if sender == 'Visualization':
                    # Visual changes Video X, width, height
                    videoDims = (X + 1, videoDims[1], videoDims[2] + (videoDims[0] - X - 1), YUpper - videoDims[1])
                    # Visual changes Transcript Y, height
                    transcriptDims = (transcriptDims[0], YUpper + 1, transcriptDims[2], transcriptDims[3] + (transcriptDims[1] - YUpper - 1))
                    # Visual changes Data Y, height
                    dataDims = (dataDims[0], YUpper + 1, dataDims[2], dataDims[3] + (dataDims[1] - YUpper - 1))
                # If the Video window has been changed ...
                elif sender == 'Video':
                    # Video changes Visual width and height
                    visualDims = (visualDims[0], visualDims[1], X - visualDims[0] - 1, YUpper - visualDims[1])
                    # Video changes Transcript Y and height
                    transcriptDims = (transcriptDims[0], YUpper + 1, transcriptDims[2], transcriptDims[3] + (transcriptDims[1] - YUpper))
                    # Video changes Data Y, height
                    dataDims = (dataDims[0], YUpper + 1, dataDims[2], dataDims[3] + (dataDims[1] - YUpper))
                # If the Transcript window has been changed ...
                elif sender == 'Transcript':
                    # Transcript changes Visual height
                    visualDims = (visualDims[0], visualDims[1], visualDims[2], YUpper - visualDims[1])
                    # Transcript changes Video height
                    videoDims = (videoDims[0], videoDims[1], videoDims[2], YUpper - videoDims[1])
                    # Transcript changes Data X, Y, width, height
                    dataDims = (X + 1, YUpper + 1, dataDims[2] + (dataDims[0] - X - 1), dataDims[3] + (dataDims[1] - YUpper - 1))
                # If the Data window has been changed ...
                elif sender == 'Data':
                    # Data changes Visual height
                    visualDims = (visualDims[0], visualDims[1], visualDims[2], YLower - visualDims[1])
                    # Data changes Transcript Y, width, height
                    transcriptDims = (transcriptDims[0], YLower + 1, X - transcriptDims[0], transcriptDims[3] + (transcriptDims[1] - YLower - 1))
                    # Data changes Video height
                    videoDims = (videoDims[0], videoDims[1], videoDims[2], YLower - videoDims[1])

                # We need to signal that we are resizing everything to reduce redundant OnSize calls
                TransanaGlobal.resizingAll = True
                # Vertical Adjustments
                # Adjust Visualization Window
                if sender != 'Visualization':
                    self.VisualizationWindow.SetDims(visualDims[0], visualDims[1], visualDims[2], visualDims[3])
                # Adjust Transcript Window
                if sender != 'Transcript':
                    self.TranscriptWindow.SetDims(transcriptDims[0], transcriptDims[1], transcriptDims[2], transcriptDims[3])
                # Adjust Video Window
                if sender != 'Video':
                    self.VideoWindow.SetDims(videoDims[0], videoDims[1], videoDims[2], videoDims[3])
                # Adjust Data Window
                if sender != 'Data':
                    self.DataWindow.SetDims(dataDims[0], dataDims[1], dataDims[2], dataDims[3])
                # We're done resizing all windows and need to re-enable OnSize calls that would otherwise be redundant
                TransanaGlobal.resizingAll = False
            # If we have only a single video window ...
            else:

                # We need to signal that we are resizing everything to reduce redundant OnSize calls
                TransanaGlobal.resizingAll = True

                if DEBUG:
                    print "ControlObjectClass.UpdateWindowPositions(2):", sender, X, YUpper, YLower
                
                # Adjust Visualization Window
                if sender != 'Visualization':
                    (wleft, wtop, wwidth, wheight) = self.VisualizationWindow.GetDimensions()
                    (oleft, otop, owidth, oheight) = (wleft, wtop, wwidth, wheight)


                    if DEBUG:
                        print "Visualization:", wleft, wtop, wwidth, wheight

                    if adjustX < wleft - 1:
                        wleft += adjustX

                    if (oleft != wleft) or (otop != wtop) or (owidth != X - wleft) or (oheight != YUpper - wtop):
                        self.VisualizationWindow.SetDims(wleft, wtop, X - wleft, YUpper - wtop)

                # Adjust Video Window
                if sender != 'Video':
                    (wleft, wtop, wwidth, wheight) = self.VideoWindow.GetDimensions()
                    (oleft, otop, owidth, oheight) = (wleft, wtop, wwidth, wheight)

                    if DEBUG:
                        print 'Video:', wleft, wtop, wwidth, wheight, X, YLower

                    # See if Secondary Monitor adjustment is needed ...
                    if wwidth + wleft - X - 1 < 0:
                        # ... we can adjust the left position
                        wleft += adjustX

                    if (oleft != X + 1) or (otop != wtop) or (owidth != wwidth + wleft - X - 1) or (oheight != YLower - wtop):
                        self.VideoWindow.SetDims(X + 1, wtop, wwidth + wleft - (X + 1), YLower - wtop)

                    self.VideoWindow.Refresh()

                # Adjust Transcript Window
                if sender != 'Transcript':
                    (wleft, wtop, wwidth, wheight) = self.TranscriptWindow.GetDimensions()
                    (oleft, otop, owidth, oheight) = (wleft, wtop, wwidth, wheight)

                    if DEBUG:
                        print "Transcript:", wleft, wtop, wwidth, wheight

                    # See if Secondary Monitor adjustment is needed ...
                    if (wheight + wtop - YUpper - 1 < 0) or ((wheight + wtop - YUpper - 1 > adjustH) and (adjustY < 0)):
                        # ... we can adjust the top position
                        wtop += adjustY

                    if YUpper + 1 < adjustY:
                        tmpAdjustedY = YUpper + adjustY + 1
                    else:
                        tmpAdjustedY = YUpper + 1

                    # Transcripts are ending up too large on OS X.  I can't figure out why, so let's just correct it here.
                    # If the bottom of the Data Window is smaller than the bottom of the Transcript Window ...
                    if (self.DataWindow.GetDimensions()[1] + self.DataWindow.GetDimensions()[3]) < \
                       (self.TranscriptWindow.GetDimensions()[1] + self.TranscriptWindow.GetDimensions()[3]):
                        # ... reduce the wheight by the difference!
                        wheight -= (self.TranscriptWindow.GetDimensions()[1] + self.TranscriptWindow.GetDimensions()[3]) - \
                                        (self.DataWindow.GetDimensions()[1] + self.DataWindow.GetDimensions()[3])

                    if (oleft != wleft) or (otop != tmpAdjustedY) or (owidth != X - wleft) or (oheight != wheight + wtop - tmpAdjustedY):
                        self.TranscriptWindow.SetDims(wleft, tmpAdjustedY, X - wleft, wheight + wtop - tmpAdjustedY)

                # Adjust Data Window
                if sender != 'Data':
                    (wleft, wtop, wwidth, wheight) = self.DataWindow.GetDimensions()
                    (oleft, otop, owidth, oheight) = (wleft, wtop, wwidth, wheight)

                    if DEBUG:
                        print "Data:", wleft, wtop, wwidth, wheight
                    
                    # See if Secondary Monitor adjustment is needed ...
                    if (wheight + wtop - YUpper - 1 < 0) or ((wheight + wtop - YUpper - 1 > adjustH) and (adjustY < 0)):
                        # ... we can adjust the top position
                        wtop += adjustY

                    if YLower + 1 < adjustY:
                        tmpAdjustedY = YLower + adjustY + 1 

                    else:
                        tmpAdjustedY = YLower + 1

                    if (oleft != X + 1) or (otop != YLower + 1) or (owidth != wwidth + wleft - X - 1) or (oheight != wheight + wtop - tmpAdjustedY):
                        self.DataWindow.SetDims(X + 1, YLower + 1, wwidth + wleft - (X + 1), wheight + wtop - tmpAdjustedY)

                if DEBUG:
                    st = "Final Window Sizes:\n"
                    st += "  Monitor %s:\t%s\n\n" % (TransanaGlobal.configData.primaryScreen, (adjustX, adjustY, adjustW, adjustH),)
                    st += "  menu:\t\t%s\n" % self.MenuWindow.GetRect()
                    st += "  visual:\t%s\n" % (self.VisualizationWindow.GetDimensions(),)
                    st += "  video:\t%s\n" % (self.VideoWindow.GetDimensions(), )
                    st += "  trans:\t%s\n" % (self.TranscriptWindow.GetDimensions(), )
                    st += "  data:\t\t%s\n" % (self.DataWindow.GetDimensions(), )

                    print
                    print st
                    print
                    print "TranWindow: Y = %d, H = %d, Total = %d" % (self.TranscriptWindow.GetDimensions()[1],  self.TranscriptWindow.GetDimensions()[3], self.TranscriptWindow.GetDimensions()[1] + self.TranscriptWindow.GetDimensions()[3])
                    print "DataWindow: Y = %d, H = %d, Total = %d" % (self.DataWindow.GetDimensions()[1],  self.DataWindow.GetDimensions()[3], self.DataWindow.GetDimensions()[1] + self.DataWindow.GetDimensions()[3])
                    print

                # We're done resizing all windows and need to re-enable OnSize calls that would otherwise be redundant
                TransanaGlobal.resizingAll = False

#        self.DataWindow.Refresh()
                    

    def VideoSizeChange(self):
        """ Signal that the Video Size has been changed via the Options > Video menu """
        # If there is a media files loaded ...
        if self.currentObj != None:
            # ... resize the video window.  This will trigger changes in all the other windows as appropriate.
            self.VideoWindow.OnSizeChange()

    def SaveTranscript(self, prompt=0, cleardoc=0, transcriptToSave=-1, continueEditing=True):
        """Save the Transcript to the database if modified.

           If prompt=1, prompt the user to confirm the save.
           If cleardoc=1, then the transcript will be cleared if the user chooses to not save.
           transcriptToSave indicates which of multiple transcripts should be saved, or -1 to save the active transcript.
           continueEditing is only applicable for Partial Transcript Editing.

           Return 1 if Transcript was saved or unchanged, and 0 if user chose to discard changes.  """

        # NOTE:  When the user presses their response to dlg below, it can shift the focus if there are multiple
        #        transcript windows open!  Therefore, remember which transcript we're working on now.
        if transcriptToSave == -1:
            transcriptToSave = self.activeTranscript

        # Figure out the correct Transcript to save!            
        panel = self.TranscriptWindow.nb.GetCurrentPage().GetChildren()[transcriptToSave]
        editor = panel.editor

        # Was the document modified?
        if self.TranscriptWindow.TranscriptModified():
            result = wx.ID_YES

            # If we should prompt the user about saving ...
            if prompt:
                # if we have a Transcript Object loaded ...
                if isinstance(editor.TranscriptObj, Transcript.Transcript):
                    # ... if we have a CLIP transcript ...
                    if editor.TranscriptObj.clip_num > 0:
                        # ... create a Clip Transcript prompt
                        pmpt = _("The Clip Transcript has changed.\nDo you want to save it before continuing?")
                    # ... if we have an Episode transcript ...
                    else:
                        # ... create an Episode Transcript prompt
                        pmpt = unicode(_('Transcript "%s" has changed.\nDo you want to save it before continuing?'), 'utf8')
                        pmpt = pmpt % editor.TranscriptObj.id
                # if we have a Document Object loaded ...
                elif isinstance(editor.TranscriptObj, Document.Document):
                    # ... create a Document prompt
                    pmpt = unicode(_('Document "%s" has changed.\nDo you want to save it before continuing?'), 'utf8')
                    pmpt = pmpt % editor.TranscriptObj.id
                # if we have a Quote Object loaded ...
                elif isinstance(editor.TranscriptObj, Quote.Quote):
                    # ... create a Quote prompt
                    pmpt = unicode(_('Quote "%s" has changed.\nDo you want to save it before continuing?'), 'utf8')
                    pmpt = pmpt % editor.TranscriptObj.id
        
                # Display the prompt and get user feedback
                dlg = Dialogs.QuestionDialog(None, pmpt, _("Question"))
                result = dlg.LocalShowModal()
                dlg.Destroy()
            # Set the active transcript to the transcript number for the transcript being saved.
            self.activeTranscript = transcriptToSave

            # If the user said to save (or was not asked) ...
            if result == wx.ID_YES:
                try:
                    # If we're using Long Transcript Editing ...
                    if TransanaConstants.partialTranscriptEdit:
                        # ... save the transcript changes 
                        editor.save_transcript(continueEditing=continueEditing)
                    # If we're NOT using Long Transcripts ...
                    else:
                        # ... save the transcript changes
                        editor.save_transcript()
                    return 1
                # If an exception is raised, report it to the user
                except TransanaExceptions.SaveError, e:
                    dlg = Dialogs.ErrorDialog(None, e.reason)
                    dlg.ShowModal()
                    dlg.Destroy()
                    return 1
            # If the user said NOT to save ...
            else:
                # ... discard them to avoid duplicate SAVE prompts
                editor.DiscardEdits()
#                if cleardoc:
#                    panel.DeletePanel()
                return 0
        # If the transcript has NOT been changed since the last save ...
        else:
            # ... and we are NOT going to continue editing it ...
            if TransanaConstants.partialTranscriptEdit and (not continueEditing):
                # ... then we need to update the contents of the editor control, restoring missing transcript lines.
                editor.UpdateCurrentContents('LeaveEditMode')
        return 1

    def SaveTranscriptAs(self):
        """Export the Transcript to an RTF file."""
        # Prompt the user to save before exporting
        if self.SaveTranscript(prompt=1, continueEditing=TransanaConstants.partialTranscriptEdit):
            # If we're using a Right-To-Left language ...
            if TransanaGlobal.configData.LayoutDirection == wx.Layout_RightToLeft:
                # ... we can only export to XML format
                wildcard = _("XML Format (*.xml)|*.xml")
            # ... whereas with Left-to-Right languages
            else:
                # ... we can export both RTF and XML formats
                wildcard = _("Rich Text Format (*.rtf)|*.rtf|XML Format (*.xml)|*.xml")
            dlg = wx.FileDialog(None, defaultDir=self.defaultExportDir,
                                wildcard=wildcard, style=wx.SAVE)
            if dlg.ShowModal() == wx.ID_OK:
                # The Default Export Directory should use the last-used value for the session but reset to the
                # video root between sessions.
                self.defaultExportDir = dlg.GetDirectory()
                fname = dlg.GetPath()
                # Mac doesn't automatically append the file extension.  Do it if necessary.
                if (TransanaGlobal.configData.LayoutDirection != wx.Layout_RightToLeft) and \
                   (dlg.GetFilterIndex() == 0) and \
                   (not fname.upper().endswith(".RTF")):
                    fname += '.rtf'
                elif (dlg.GetFilterIndex() == 1) and (not fname.upper().endswith(".XML")):
                    fname += '.xml'
                if os.path.exists(fname):
                    prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                    dlg2 = Dialogs.QuestionDialog(None, prompt % fname, _('Transana Confirmation'))
                    TransanaGlobal.CenterOnPrimary(dlg2)
                    if dlg2.LocalShowModal() == wx.ID_YES:
                        self.TranscriptWindow.SaveTranscriptAs(fname)
                    dlg2.Destroy()
                else:
                    self.TranscriptWindow.SaveTranscriptAs(fname)
            dlg.Destroy()
        # If the user doesn't SAVE ...
        else:
            # ... mark that the transcript IS changed so we don't lose these edits!
            self.TranscriptWindow[self.activeTranscript].dlg.editor.MarkDirty()

    def PropagateChanges(self, transcriptWindowNumber):
        """ Propagate changes in an Episode transcript down to derived clips """
        # First, let's save the changes in the Transcript.  We don't want to propagate changes, then end up
        # not saving them in the source!
        if TransanaConstants.partialTranscriptEdit:
            saveResult = self.SaveTranscript(prompt=1, continueEditing=True)
        else:
            saveResult = self.SaveTranscript(prompt=1)
        if saveResult:
            if isinstance(self.currentObj, Document.Document):
                # Start up the Propagate Episode Transcript Changes tool
                propagateDlg = PropagateChanges.PropagateObjectChanges(self, 'Document')

            # If we are working with an Episode Transcript ...
            elif type(self.currentObj).__name__ == 'Episode':
                # Start up the Propagate Episode Transcript Changes tool
                propagateDlg = PropagateChanges.PropagateObjectChanges(self, 'Episode')

            elif isinstance(self.currentObj, Quote.Quote):

                # If the user has updated the Quote's Keywords, self.currentObj will NOT reflect this.
                # Therefore, we need to load a new copy of the Quote to get the latest keywords for propagation.
                tempObj = Quote.Quote(num = self.currentObj.number)
                if TransanaConstants.USESRTC:
                    # Start up the Propagate "Clip" Changes tool
                    propagateDlg = PropagateChanges.PropagateClipChanges(self.MenuWindow,
                                                                         'Quote',
                                                                         self.currentObj,
                                                                         -1,
                                                                         self.TranscriptWindow.dlg.editor.GetFormattedSelection('XML'),
                                                                         newKeywordList=tempObj.keyword_list)
                else:
                    # Start up the Propagate "Clip" Changes tool
                    propagateDlg = PropagateChanges.PropagateClipChanges(self.MenuWindow,
                                                                         'Quote',
                                                                         self.currentObj,
                                                                         -1,
                                                                         self.TranscriptWindow.dlg.editor.GetRTFBuffer(),
                                                                         newKeywordList=tempQuote.keyword_list)

            # If we are working with a Clip Transcript ...
            elif type(self.currentObj).__name__ == 'Clip':

                # If the user has updated the clip's Keywords, self.currentObj will NOT reflect this.
                # Therefore, we need to load a new copy of the clip to get the latest keywords for propagation.
                tempClip = Clip.Clip(self.currentObj.number)
                if TransanaConstants.USESRTC:
                    # Start up the Propagate Clip Changes tool
                    propagateDlg = PropagateChanges.PropagateClipChanges(self.MenuWindow,
                                                                         'Clip',
                                                                         self.currentObj,
                                                                         transcriptWindowNumber,
                                                                         self.TranscriptWindow.dlg.editor.GetFormattedSelection('XML'),
                                                                         newKeywordList=tempClip.keyword_list)
                else:
                    # Start up the Propagate Clip Changes tool
                    propagateDlg = PropagateChanges.PropagateClipChanges(self.MenuWindow,
                                                                         'Clip',
                                                                         self.currentObj,
                                                                         transcriptWindowNumber,
                                                                         self.TranscriptWindow.dlg.editor.GetRTFBuffer(),
                                                                         newKeywordList=tempClip.keyword_list)

        # If the user chooses NOT to save the Transcript changes ...
        else:
            # ... let them know that nothing was propagated!
            dlg = Dialogs.InfoDialog(None, _("You must save the transcript if you want to propagate the changes."))
            dlg.ShowModal()
            dlg.Destroy()

    def PropagateObjectKeywords(self, objType, objNum, newKeywordList):
        """ When Document or Episode Keywords are added, this will allow the user to propagate new keywords to all
            Quotes or Clips created from that object if desired. """
        if objType == _('Document'):
            tmpObj = Document.Document(objNum)
        elif objType == _('Episode'):
            # Get the Episode Object
            tmpObj = Episode.Episode(objNum)
        # Initialize a list of keywords that have been added to the Document / Episode
        keywordsToAdd = []
        # Iterate through the NEW Keywords list
        for kw in newKeywordList:
            # Initialize that the new keyword has NOT been found
            found = False
            # Iterate through the OLD Keyword list (The "in" operator doesn't work here.)
            for kw2 in tmpObj.keyword_list:
                # See if the new Keyword matches the old Keyword
                if (kw.keywordGroup == kw2.keywordGroup) and (kw.keyword == kw2.keyword):
                    # If so, flag it as found ...
                    found = True
                    # ... and stop iterating
                    break
            # If the NEW Keyword was NOT found ...
            if not found:
                # ... add it to the list of keywords to add to Quotes / Clips
                keywordsToAdd.append(kw)

        if objType == _('Document'):
            # Get the list of Quotes created from this Episode
            childList = DBInterface.list_of_quotes_by_document(objNum)
        elif objType == _('Episode'):
            # Get the list of clips created from this Episode
            childList = DBInterface.list_of_clips_by_episode(objNum)
        # If there are children that have been created from this object AND Keywords have been added to the object ...
        if (len(childList) > 0) and (len(keywordsToAdd) > 0):
            # If there's only one Keyword to Add, prompt for it individually.
            # (This happens when multiple keywords are dragged to a Transcript with Quick Clips mode OFF!)
            if len(keywordsToAdd) == 1:
                if objType == _('Document'):
                    prompt = unicode(_("Do you want to add keyword %s : %s to all Quotes created from Document %s?"), 'utf8')
                elif objType == _('Episode'):
                    prompt = unicode(_("Do you want to add keyword %s : %s to all Clips created from Episode %s?"), 'utf8')
                data = (keywordsToAdd[0].keywordGroup, keywordsToAdd[0].keyword, tmpObj.id)
            else:
                if objType == _('Document'):
                    prompt = unicode(_("Do you want to add the new keywords to all Quotes created from Episode %s?"), 'utf8')
                elif objType == _('Episode'):
                    prompt = unicode(_("Do you want to add the new keywords to all Clips created from Episode %s?"), 'utf8')
                data = (tmpObj.id,)
            # ... build a dialog to prompt the user about adding them to the children
            tmpDlg = Dialogs.QuestionDialog(None, prompt % data,
                                            unicode(_("%s Keyword Propagation"), 'utf8') % objType, noDefault = True)
            # Prompt the user.  If the user says YES ...
            if tmpDlg.LocalShowModal() == wx.ID_YES:
                # Create a Progress Dialog
                if objType == _('Document'):
                    title = _('Document Keyword Propagation')
                    prompt = unicode(_("Adding keywords to Document %s"), 'utf8')
                elif objType == _('Episode'):
                    title = _('Episode Keyword Propagation')
                    prompt = unicode(_("Adding keywords to Episode %s"), 'utf8')
                progress = wx.ProgressDialog(title, prompt % tmpObj.id, parent=self.TranscriptWindow, style = wx.PD_APP_MODAL | wx.PD_AUTO_HIDE)
                progress.Centre()
                # That's not working.  Let's try this ...
                TransanaGlobal.CenterOnPrimary(progress)
                # Initialize the Object Counter for the progess dialog
                objCount = 0.0
                # Iterate through the Child list 
                for childRec in childList:
                    # increment the child counter for the progress dialog
                    objCount += 1.0
                    if objType == _('Document'):
                        # Load the Quote.
                        tmpChildObj = Quote.Quote(childRec['QuoteNum'])
                        childType = 'Quote'
                    elif objType == _('Episode'):
                        # Load the Clip.
                        tmpChildObj = Clip.Clip(childRec['ClipNum'])
                        childType = 'Clip'
                    # Start Exception Handling
                    try:
                        # Lock the Child
                        tmpChildObj.lock_record()
                        # Add the new Keywords
                        for kw in keywordsToAdd:
                            tmpChildObj.add_keyword(kw.keywordGroup, kw.keyword)
                        # Save the Child
                        tmpChildObj.db_save()
                        # Unlock the Child
                        tmpChildObj.unlock_record()

                        if not TransanaConstants.singleUserVersion:
                            if TransanaGlobal.chatWindow != None:
                                # We need to update the Keyword Visualization for the current ClipObject
                                if objType == _('Document'):

                                    if DEBUG:
                                        print 'Message to send = "UKV %s %s %s"' % (childType, tmpChildObj.number, tmpChildObj.source_document_num)

                                    TransanaGlobal.chatWindow.SendMessage("UKL %s %s" % (childType, tmpChildObj.number))
                                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % (childType, tmpChildObj.number, tmpChildObj.source_document_num))
                                elif objType == _('Episode'):

                                    if DEBUG:
                                        print 'Message to send = "UKV %s %s %s"' % (childType, tmpChildObj.number, tmpChildObj.episode_num)
                                
                                    TransanaGlobal.chatWindow.SendMessage("UKL %s %s" % (childType, tmpChildObj.number))
                                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % (childType, tmpChildObj.number, tmpChildObj.episode_num))

                        # Increment the Progress Dialog
                        progress.Update(int((objCount / len(childList) * 100)))
                    # Handle Exceptions
                    except TransanaExceptions.RecordLockedError, e:
                        if objType == _('Document'):
                            prompt = unicode(_('New keywords were not added to Quote "%s"\nin Document "%s"\nbecause the Quote record was locked by %s.'), 'utf8')
                        elif objType == _('Episode'):
                            prompt = unicode(_('New keywords were not added to Clip "%s"\nin Collection "%s"\nbecause the Clip record was locked by %s.'), 'utf8')
                        errDlg = Dialogs.ErrorDialog(self.MenuWindow, prompt % (tmpChildObj.id, tmpChildObj.GetNodeString(False), e.user))
                        errDlg.ShowModal()
                        errDlg.Destroy()
                # Destroy the Progress Dialog
                progress.Destroy()

                # Need to Update the Keyword Visualization
                self.UpdateKeywordVisualization()

                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    if TransanaGlobal.chatWindow != None:
                        if objType == _('Document'):
                            childType = 'Document'
                        elif objType == _('Episode'):
                            childType = 'Episode'
                        
                        # We need to update the Episode Keyword Visualization
                        if DEBUG:
                            print 'Message to send = "UKV %s %s %s"' % (childType, objNum, 0)

                        TransanaGlobal.chatWindow.SendMessage("UKL %s %s" % (childType, objNum))
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % (childType, objNum, 0))

            # Destroy the User Prompt dialog
            tmpDlg.Destroy()

    def MultiSelect(self, transcriptWindowNumber):
        """ Make selections in all other transcripts to match the selection in the identified transcript """
        # Determine the start and end times of the selection in the identified transcript window
        (start, end) = self.TranscriptWindow.dlg.editor.get_selected_time_range()
        # If start is 0 ...
        if start == 0:
            # ... and we're in a Clip ...
            if isinstance(self.currentObj, Clip.Clip):
                # ... then the start should be the Clip Start
                start = self.currentObj.clip_start
        # If there is not following time code ...
        if end <= 0:
            # ... and we're in a Clip ...
            if isinstance(self.currentObj, Clip.Clip):
                # ... use the Clip Stop value ...
                end = self.currentObj.clip_stop
            # ... otherwise ...
            else:
                # ... use the length of the media file
                end = self.GetMediaLength(entire = True)

        # Iterate through all Transcript Windows
        for trWin in self.TranscriptWindow.nb.GetCurrentPage().GetChildren():
            # If we have a transcript window other than the identified one ...
            if trWin.panelNum != transcriptWindowNumber:
                # ... highlight the full text of the video selection
                # If I don't use a time ever so slightly earlier than start, the first time-coded segment of the
                # selection is left out!!
                trWin.editor.scroll_to_time(start - 2)
                trWin.editor.select_find(str(end))
                # Check for time codes at the selection boundaries
                trWin.editor.CheckTimeCodesAtSelectionBoundaries()
                # Once selections are set (later), update the Selection Text
                wx.CallLater(200, self.UpdateSelectionTextLater, trWin.panelNum)

    def MultiPlay(self):
        """ Play the current video based on selections in multiple transcripts """
        # Save the cursors for all transcripts (!)
        self.SaveAllTranscriptCursors()
        # Get the Transcript Selection information from all transcript windows.
        transcriptSelectionInfo = self.GetMultipleTranscriptSelectionInfo()
        # Initialize the clip start time to the end of the media file
        earliestStartTime = self.GetMediaLength(True)
        # Initialise the clip end time to the beginning of the media file
        latestEndTime = 0
        # Iterate through the Transcript Selection Info gathered above
        for (transcriptNum, startTime, endTime, text) in transcriptSelectionInfo:
            # If the transcript HAS a selection ...
            if text != "":
                # Check to see if this transcript starts before our current earliest start time, but only if
                # if actually contains text.
                if (startTime < earliestStartTime) and (text != ''):
                    # If so, this is our new earliest start time.
                    earliestStartTime = startTime
                # Check to see if this transcript ends after our current latest end time, but only if
                # if actually contains text.
                if (endTime > latestEndTime) and (text != ''):
                    # If so, this is our new latest end time.
                    latestEndTime = endTime
        # Set the Video Selection to the boundary times
        self.SetVideoSelection(earliestStartTime, latestEndTime)
        # Play the video selection!
        self.Play()

    def UpdateDataWindow(self):
        """ Update the Data Window, as when the "Update Database Window" command is issued """
        # NOTE:  This is called in MU when one user imports a database while another user is connected.
        # Tell the Data Window's Database Tree Tab's Tree to refresh itself
        self.DataWindow.DBTab.tree.refresh_tree()

    def UpdateDataWindowOnDocumentEdit(self):
        """ Update the DataWindow because a Document was edited """
        self.DataWindow.UpdateDataWindow()

    def GetQuoteDataForDocument(self, documentNum, textPos=-1, textSel=(-2, -2)):
        """ We need the Quote Data for a Document.  This data COULD be different in the live document, if edited,
            than in the database.  This method checks to see if the Document in question is open.  If it is, it
            returns the live Quote data.  If not, it returns the data from the database. """

        # Initialize a variable to hold the document when we find it
        docFound = None
        # For each Page in the TranscriptWindow Notebook ...
        for tabNum in range(self.TranscriptWindow.nb.GetPageCount()):
            for panelNum in range(len(self.TranscriptWindow.nb.GetPage(tabNum).GetChildren())):
                # ... if the text contained is a Document AND that document is the document we're looking for ...
                if isinstance(self.TranscriptWindow.nb.GetPage(tabNum).GetChildren()[panelNum].editor.TranscriptObj, Document.Document) and \
                   (self.TranscriptWindow.nb.GetPage(tabNum).GetChildren()[panelNum].editor.TranscriptObj.number == documentNum):
                    # ... then select THIS document ...
                    docFound = self.TranscriptWindow.nb.GetPage(tabNum).GetChildren()[panelNum].editor.TranscriptObj
                    # ... and stop looking
                    break
                # If we found the document ...
                if docFound != None:
                    # ... we really can stop looking!
                    break

        # Get the Quote Data from the Database
        data = DBInterface.list_of_quotes_by_document(documentNum, textPos, textSel)

        # Iterate through the list of quotes found in the DATABASE.  We need to go backwards because we may
        # delete some elements, which screws up list iteration otherwise
        for x in range(len(data)-1, -1, -1):
            # If the Quote is in the LIVE document ...
            if data[x]['QuoteNum'] in docFound.quote_dict.keys():
                # ... use the live document's quote positions
                (data[x]['StartChar'], data[x]['EndChar']) = docFound.quote_dict[data[x]['QuoteNum']]
            # If the quote has been deleted from the LIVE Document ...
            else:
                # ... drop it from the Database List
                del(data[x])
        
        return data

    def DataWindowHasSearchNodes(self):
        """ Returns the number of Search Nodes in the DataWindow's Database Tree """
        # Find the Search Node, using the localized label
        searchNode = self.DataWindow.DBTab.tree.select_Node((_('Search'),), 'SearchRootNode')
        # If there's not a problem with the node ...
        if (searchNode != None) and searchNode.IsOk():
            # ... return information about whether it has child nodes
            return self.DataWindow.DBTab.tree.ItemHasChildren(searchNode)
        # If there is a problem with the node ...
        else:
            # ... just return False.  (I'm not sure why this fails occasionally.)
            return False

    def RemoveDataWindowKeywordExamples(self, keywordGroup, keyword, clipNum):
        """ Remove Keyword Examples from the Data Window """
        # First, remove the Keyword Example from the Database Tree
        # Load the specified Clip record.  Skip the Transcript to speed the process up.
        tempClip = Clip.Clip(clipNum, skipText=True)
        # Prepare the Node List for removing the Keyword Example Node
        nodeList = (_('Keywords'), keywordGroup, keyword, tempClip.id)
        # Call the DB Tree's delete_Node method.  Include the Clip Record Number so the correct Clip entry will be removed.
        self.DataWindow.DBTab.tree.delete_Node(nodeList, 'KeywordExampleNode', tempClip.number)

    def UpdateDataWindowKeywordsTab(self):
        """ Update the Keywords Tab in the Data Window """
        # If the Keywords Tab is the currently displayed tab ...
        if self.DataWindow.nb.GetPageText(self.DataWindow.nb.GetSelection()) == unicode(_('Keywords'), 'utf8'):
            # ... then refresh the Tab
            self.DataWindow.KeywordsTab.Refresh()

    def UpdateSSLStatus(self, sslValue):
        """ Update the SSL Status of Transana """
        self.DataWindow.UpdateSSLStatus(sslValue)

    def CreateTranscriptlessClip(self):
        """ Trigger the creation of a  Clip without a transcript from outside of the Daabase Tree.
            If a single Collection is selected in the Database Tree, a Standard Clip will be created.
            If one or more Keyword records are selected in the Database Tree, a Quick Clip will be created. """
        # Get the list of selected Nodes in the Database Tree
        dbTreeSelections = self.DataWindow.DBTab.GetSelectedNodeInfo()

        # If the selection list has ONE Collection, we are creating a Transcript-less Standard Clip
        if (len(dbTreeSelections) == 1) and (dbTreeSelections[0][3] == 'CollectionNode'):

            # We also need to know the number of the original Transcript Record
            if self.TranscriptWindow.dlg.editor.TranscriptObj.clip_num == 0:
                # If we have an Episode Transcript, we need the Transcript Number
                transcriptNum = self.TranscriptWindow.dlg.editor.TranscriptObj.number
            else:
                # If we have a Clip Transcript, we need the original Transcript Number, not the Clip Transcript Number.
                # We can get that from the ControlObject's "currentObj", which in this case will be the Clip!
                transcriptNum = self.currentObj.transcripts[self.activeTranscript].source_transcript

            # If our source is an Episode ...
            if isinstance(self.currentObj, Episode.Episode):
                # ... we can just use the ControlObject's currentObj's object number
                episodeNum = self.currentObj.number
            # If our source is a Clip ...
            elif isinstance(self.currentObj, Clip.Clip):
                # ... we need the ControlObject's currentObj's originating episode number
                episodeNum = self.currentObj.episode_num

            # The Clip's Start and End times can be obtained from the Video Start Point and Video End Point,
            # which were set by creating the selection in the Waveform.
            startTime = self.GetVideoStartPoint()
            endTime = self.GetVideoEndPoint()
            # Since this is by definition transcript-less, we won't have a transcript.  However, we use this
            # to signal that we are intentionally leaving the transcript blank.
            text = u'<(transcript-less clip)>'

            # We now have enough information to populate a ClipDragDropData object to pass to the Clip Creation method.
            clipData = DragAndDropObjects.ClipDragDropData(transcriptNum, episodeNum, startTime, endTime, text, text, videoCheckboxData=self.GetVideoCheckboxDataForClips(startTime))

            # let's convert that object into a portable string using cPickle. (cPickle is faster than Pickle.)
            pdata = cPickle.dumps(clipData, 1)
            # Create a CustomDataObject with the format of the ClipDragDropData Object
            cdo = wx.CustomDataObject(wx.CustomDataFormat("ClipDragDropData"))
            # Put the pickled data object in the wxCustomDataObject
            cdo.SetData(pdata)

            # Open the Clipboard
            wx.TheClipboard.Open()
            # ... then copy the data to the clipboard!
            wx.TheClipboard.SetData(cdo)
            # Close the Clipboard
            wx.TheClipboard.Close()

            # Now we can create the Standard Clip by triggering the DataWindow's add_clip method as if we'd dropped a
            # selection on the selected Collection
            self.DataWindow.DBTab.add_clip(dbTreeSelections[0][1])

        # If we have one or more selections that are Keywords ...
        elif (len(dbTreeSelections) > 0) and (dbTreeSelections[0][3] == 'KeywordNode'):
            # ... we can create a Quick Clip, using "transcriptless" mode
            self.CreateQuickClip(transcriptless=True)
        # If neither of these conditions applies ...
        else:
            # ... create an error message
            msg = _("You must select one Collection in the Data Tree to create a Transcript-less Standard Clip,") + \
                  '\n' + _("or select one or more Keywords in the Data Tree to create a Transcript-less Quick Clip.")
            msg = unicode(msg, 'utf8')
            # Display the error message and then clean up.
            dlg = Dialogs.ErrorDialog(None, msg)
            dlg.ShowModal()
            dlg.Destroy()

    def CreateQuickClip(self, transcriptless=False):
        """ Trigger the creation of a Quick Clip from outside of the Database Tree.
            The "transcriptless" parameter will be True if triggered from the Visualization Window, but will be
            left off otherwise. """

        # Get the list of selected Nodes in the Database Tree
        dbTreeSelections = self.DataWindow.DBTab.GetSelectedNodeInfo()

        # The selection list must not be empty , and Keywords MUST be selected, or we don't know what keyword to base the Quick Clip on
        if (len(dbTreeSelections) > 0) and (dbTreeSelections[0][3] == 'KeywordNode'):
            if not transcriptless:
                # Get the Transcript Selection information from the ControlObject, since we can't communicate with the
                # TranscriptEditor directly.
                (transcriptNum, startTime, endTime, text) = self.GetTranscriptSelectionInfo()
                # Initialize the Episode Number to 0
                episodeNum = 0
                # If our source is an Episode ...
                if isinstance(self.currentObj, Episode.Episode):
                    # ... we can just use the ControlObject's currentObj's object number
                    episodeNum = self.currentObj.number
                    # If we are at the end of a transcript and there are no later time codes, Stop Time will be -1.
                    # This is, of course, incorrect, and we must replace it with the Episode Length.
                    if endTime <= 0:
                        endTime = self.VideoWindow.GetMediaLength()
                # If our source is a Clip ...
                elif isinstance(self.currentObj, Clip.Clip):
                    # ... we need the ControlObject's currentObj's originating episode number
                    episodeNum = self.currentObj.episode_num
                    # Sometimes with a clip, we get a startTime of 0 from the TranscriptSelectionInfo() method.
                    # This is, of course, incorrect, and we must replace it with the Clip Start Time.
                    if startTime == 0:
                        startTime = self.currentObj.clip_start
                    # Sometimes with a clip, we get an endTime of 0 from the TranscriptSelectionInfo() method.
                    # This is, of course, incorrect, and we must replace it with the Clip Stop Time.
                    if endTime <= 0:
                        endTime = self.currentObj.clip_stop
            else:

                # We also need to know the number of the original Transcript Record
                if self.TranscriptWindow.dlg.editor.TranscriptObj.clip_num == 0:
                    # If we have an Episode Transcript, we need the Transcript Number
                    transcriptNum = self.TranscriptWindow.dlg.editor.TranscriptObj.number
                else:
                    # If we have a Clip Transcript, we need the original Transcript Number, not the Clip Transcript Number.
                    # We can get that from the ControlObject's "currentObj", which in this case will be the Clip!
                    transcriptNum = self.currentObj.transcripts[self.activeTranscript].source_transcript

                # If our source is an Episode ...
                if isinstance(self.currentObj, Episode.Episode):
                    # ... we can just use the ControlObject's currentObj's object number
                    episodeNum = self.currentObj.number
                # If our source is a Clip ...
                elif isinstance(self.currentObj, Clip.Clip):
                    # ... we need the ControlObject's currentObj's originating episode number
                    episodeNum = self.currentObj.episode_num

                startTime = self.GetVideoStartPoint()
                endTime = self.GetVideoEndPoint()
                text = u'<(transcript-less clip)>'

            # We now have enough information to populate a ClipDragDropData object to pass to the Clip Creation method.
            clipData = DragAndDropObjects.ClipDragDropData(transcriptNum, episodeNum, startTime, endTime, text, videoCheckboxData=self.GetVideoCheckboxDataForClips(startTime))

            # Let's assemble the keyword list
            kwList = []
            # For each selected node in the DB Tree ...
            for selection in dbTreeSelections:
                # ... get the Node information
                (nodeName, nodeRecNum, nodeParent, nodeType) = selection
                # ... and add the keyword info to the Keyword List
                kwList.append((nodeParent, nodeName))
            # Pass the accumulated data to the CreateQuickClip method, which is in the DragAndDropObjects module
            # because drag and drop is an alternate way to create a Quick Clip.
            DragAndDropObjects.CreateQuickClip(clipData, kwList[0][0], kwList[0][1], self.DataWindow.DBTab.tree, extraKeywords = kwList[1:])

        # If there is something OTHER than a Keyword selected in the Database Tree ...
        elif (len(dbTreeSelections) > 0):
            # ... create an error message
            msg = unicode(_("You must select a Keyword in the Data Tree to create a Quick Clip this way."), 'utf8')
            # Display the error message and then clean up.
            dlg = Dialogs.ErrorDialog(None, msg)
            dlg.ShowModal()
            dlg.Destroy()

    def CreateQuickQuote(self):
        """ Trigger the creation of a Quick Quote from outside of the Database Tree. """

        # Get the list of selected Nodes in the Database Tree
        dbTreeSelections = self.DataWindow.DBTab.GetSelectedNodeInfo()

        # The selection list must not be empty , and Keywords MUST be selected, or we don't know what keyword to base the Quick Quote on
        if (len(dbTreeSelections) > 0) and (dbTreeSelections[0][3] == 'KeywordNode'):
            try:
                # Get the Document Selection information from the ControlObject.
                (documentNum, startChar, endChar, text) = self.GetDocumentSelectionInfo()
            except:
                if DEBUG or True:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                # Create an error message.
                msg = _('You must make a selection in a document to be able to add a Quote.')
                errordlg = Dialogs.InfoDialog(None, msg)
                errordlg.ShowModal()
                errordlg.Destroy()
            # Initialize the Source Document Number to 0
            sourceDocumentNum = 0
            # If our source is a Document ...
            if self.GetCurrentItemType() == 'Document':
                # ... we can just use the Transcript Window's CurrentObj's object number
                sourceDocumentNum = self.TranscriptWindow.GetCurrentObject().number
            # If our source is a Quote ...
            elif self.GetCurrentItemType() == 'Quote':
                # ... we need the currentObj's originating document number
                sourceDocumentNum = self.TranscriptWindow.GetCurrentObject().source_document_num
                startChar += self.TranscriptWindow.GetCurrentObject().start_char
                endChar += self.TranscriptWindow.GetCurrentObject().start_char

            # Let's assemble the keyword list
            kwList = []
            # For each selected node in the DB Tree ...
            for selection in dbTreeSelections:
                # ... get the Node information
                (nodeName, nodeRecNum, nodeParent, nodeType) = selection
                # ... and add the keyword info to the Keyword List
                kwList.append((nodeParent, nodeName))

            # We now have enough information to populate a QuoteDragDropData object to pass to the Quote Creation method.
            quoteData = DragAndDropObjects.QuoteDragDropData(documentNum, sourceDocumentNum, startChar, endChar, text)
            # Pass the accumulated data to the CreateQuickQuote method, which is in the DragAndDropObjects module
            # because drag and drop is an alternate way to create a Quick Quote.
            DragAndDropObjects.CreateQuickQuote(quoteData, kwList[0][0], kwList[0][1], self.DataWindow.DBTab.tree, extraKeywords=kwList[1:])

        # If there is something OTHER than a Keyword selected in the Database Tree ...
        elif (len(dbTreeSelections) > 0):
            # ... create an error message
            msg = unicode(_("You must select a Keyword in the Data Tree to create a Quick Quote this way."), 'utf8')
            # Display the error message and then clean up.
            dlg = Dialogs.ErrorDialog(None, msg)
            dlg.ShowModal()
            dlg.Destroy()

    def ChangeLanguages(self):
        """ Update all screen components to reflect change in the selected program language """
        self.ClearAllWindows()

        self.CloseAllImages()
        self.CloseAllReports()

##        # Let's look at the issue of database encoding.  We only need to do something if the encoding is NOT UTF-8
##        # or if we're on Windows single-user version.
##        if (TransanaGlobal.encoding != 'utf8') or \
##           (('wxMSW' in wx.PlatformInfo) and (TransanaConstants.singleUserVersion)):
##            # If it's not UTF-*, then if it is Russian, use KOI8r
##            if TransanaGlobal.configData.language == 'ru':
##                newEncoding = 'koi8_r'
##            # If it's Chinese, use the appropriate Chinese encoding
##            elif TransanaGlobal.configData.language == 'zh':
##                newEncoding = TransanaConstants.chineseEncoding
##            # If it's Eastern European Encoding, use 'iso8859_2'
##            elif TransanaGlobal.configData.language == 'easteurope':
##                newEncoding = 'iso8859_2'
##            # If it's Greek, use 'iso8859_7'
##            elif TransanaGlobal.configData.language == 'el':
##                newEncoding = 'iso8859_7'
##            # If it's Japanese, use cp932
##            elif TransanaGlobal.configData.language == 'ja':
##                newEncoding = 'cp932'
##            # If it's Korean, use cp949
##            elif TransanaGlobal.configData.language == 'ko':
##                newEncoding = 'cp949'
##            # Otherwise, fall back to UTF-8
##            else:
##                newEncoding = 'utf8'
##
##            # If we're changing encodings, we need to do a little work here!
##            if newEncoding != TransanaGlobal.encoding:
##                msg = _('Database encoding is changing.  To avoid potential data corruption, \nTransana must close your database before proceeding.')
##                tmpDlg = Dialogs.InfoDialog(None, msg)
##                tmpDlg.ShowModal()
##                tmpDlg.Destroy()
##
##                # We should get a new database.  This call will actually update our encoding if needed!
##                self.GetNewDatabase()
                
        self.MenuWindow.ChangeLanguages()
        self.VisualizationWindow.ChangeLanguages()
        self.DataWindow.ChangeLanguages()
        self.VideoWindow.ChangeLanguages()

        # Updating the Data Window automatically updates the Headers on the Video and Transcript windows!
        self.TranscriptWindow.ChangeLanguages()
        # If we're in multi-user mode ...
        if not TransanaConstants.singleUserVersion:
            # We need to update the ChatWindow too
            self.ChatWindow.ChangeLanguages()

    def AutoTimeCodeEnableTest(self):
        """ Test to see if the Fixed-Increment Time Code menu item should be enabled """
        # See if the transcript has some time codes.  If it does, we cannot enable the menu item.
        return len(self.TranscriptWindow.dlg.editor.timecodes) == 0
        
    def AutoTimeCode(self):
        """ Add fixed-interval time codes to a transcript """
        # Ask the Transcript Editor to handle AutoTimeCoding and let us know if it worked.
        # "result" indicates whether the menu item needs to be disabled!
        result = self.TranscriptWindow.dlg.editor.AutoTimeCode()
        # Return the function result obtained
        return result
        
    def AdjustIndexes(self, adjustmentAmount):
        """ Adjust Transcript Time Codes by the specified amount """
        self.TranscriptWindow.AdjustIndexes(adjustmentAmount)

    def TextTimeCodeConversion(self):
        """ Convert Text (H:MM:SS.hh) Time Codes to Transana's Format """
        # Call the Transcription UI's Text Time Code Conversion method
        self.TranscriptWindow.TextTimeCodeConversion()

    def __repr__(self):
        """ Return a string representation of information about the ControlObject """
        tempstr = "Control Object contents:\nVideoFilename = %s\nVideoStartPoint = %s\nVideoEndPoint = %s\n"  % (self.VideoFilename, self.VideoStartPoint, self.VideoEndPoint)
        tempstr += 'Current active transcript: %d\n' % (self.activeTranscript)
        return tempstr.encode('utf8')

    def _get_activeTranscript(self):
        """ "Getter" for the activeTranscript property """
        # We need to return the activeTranscript value
        return self._activeTranscript

    def _set_activeTranscript(self, transcriptNum):
        """ "Setter" for the activeTranscript property """
        # If we're not shutting down the system ... (avoids an exception when shutting down!)
        if not self.shuttingDown:
            # Initiate exception handling.  (Shutting down Transana generates exceptions here!)
            try:
                # Set the underlying data value to the new window number
                self._activeTranscript = transcriptNum
                # Set the Menus to match the active transcript's Edit state
                if not self.shuttingDown and (self.MenuWindow != None):
                    # Enable or disable the transcript menu item options
                    self.MenuWindow.SetTranscriptEditOptions(not self.TranscriptWindow.dlg.editor.get_read_only())
            except:

                if DEBUG:
                    print "Exception in ControlObjectClass._set_activeTranscript()"

                    import traceback
                    traceback.print_exc(file=sys.stdout)
                
                # If this occurs, set activeTranscript to 0
                self._activeTranscript = 0

    # define the activeTranscript property for the ControlObject.
    # Doing this as a property allows automatic labeling of the active transcript window.
    activeTranscript = property(_get_activeTranscript, _set_activeTranscript, None)

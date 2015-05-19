# Copyright (C) 2007 The Board of Regents of the University of Wisconsin System 
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

""" This module implements the Notes Browser.  It's used for browsing through Notes objects """

DEBUG = False
if DEBUG:
    print "NotesBrowser.py DEBUG is ON!!"

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx
# import Python's sys module
import sys
# import Transana's Clip object
import Clip
# import Transana's Collection object
import Collection
# import Transana's DatabaseTreeTab (for the _nodeData definition)
import DatabaseTreeTab
# import Transana's Database Interface
import DBInterface
# import Transana's Dialogs
import Dialogs
# import Transana's Episode object
import Episode
# import Transana's Note object
import Note
# import Transana's Note Editor
import NoteEditor
# Import Transana's Note Properties Form
import NotePropertiesForm
# import Notes Report Generator
import ReportGeneratorForNotes
# import Transana's Series object
import Series
# import Transana's Constants
import TransanaConstants
# import Transana's Exceptions
import TransanaExceptions
# import Transana's Global variables
import TransanaGlobal
# import Transana's Transcript object
import Transcript

class MySplitter(wx.SplitterWindow):
    """ A local subclass of the wxSplitterWindow """
    def __init__(self, parent, ID):
        """ Initialize the Splitter Window """
        # Remember your ancestors
        self.parent = parent
        # Initialize the Splitter Window
        wx.SplitterWindow.__init__(self, parent, ID, style = wx.SP_LIVE_UPDATE)

    def OnHelp(self, event):
        """ Implement the Help function (required by the _NotePanel)"""
        # If the global Menu Window exists ...
        if (TransanaGlobal.menuWindow != None):
            # ... call Help through its Control Object!  We need the Notes Browser help.
            TransanaGlobal.menuWindow.ControlObject.Help("Notes Browser")

    def OnClose(self, event):
        """ Implement the Close function (required by the _NotePanel) """
        # To close the Splitter Window, we need to close its Parent Object!
        self.parent.OnClose(event)


class NotesBrowser(wx.Dialog):
    """ The Transana Notes Browser. """
    def __init__(self,parent,id,title):
        # Create a Dialog to house the Notes Browser
        wx.Dialog.__init__(self,parent,-1, title, size = (800,600), style=wx.DEFAULT_FRAME_STYLE | wx.NO_FULL_REPAINT_ON_RESIZE)
        # Get the Transana Icon
        transanaIcon = wx.Icon("images/Transana.ico", wx.BITMAP_TYPE_ICO)
        # Specify the Transana Icon as the Dialog's Icon
        self.SetIcon(transanaIcon)
        # initialize the pointer to the Current Active Note to None
        self.activeNote = None
        # Initialize a variable to track if the note ID is changed
        self.originalNoteID = ''
        # Initialize the variable that indicates which Tree is currently active
        # (used in right-click menu methods)
        self.activeTree = None
        # Initialize the Search Text string
        self.searchText = ''
        # Define the Control Object
        self.ControlObject = None
        # Create the Notes Browser's main Sizer
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        # Create a Splitter Window
        splitter = MySplitter(self, -1)
        # Put the Splitter Window on the Main Sizer
        mainSizer.Add(splitter, 1, wx.EXPAND)
        # Create a panel for the Notes Tree, placed on the Splitter Window
        self.treePanel = wx.Panel(splitter, -1, style = wx.RAISED_BORDER)
        # Create a Notebook control, placed on the Notes Tree panel
        self.treeNotebook = wx.Notebook(self.treePanel, -1, style=wx.CLIP_CHILDREN)
        # Create a Sizer for the Note Tree panel
        noteTreeSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        # Create a Panel for the Notes tab of the Notebook
        self.treeNotebookNotesTab = wx.Panel(self.treeNotebook, -1)
        # Add the Notes Tab to the Notebook
        self.treeNotebook.AddPage(self.treeNotebookNotesTab, _("Notes"), True)
        # Create a Sizer for the Notes Tab
        noteTreeCtrlSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Create a Tree Control and place it on the Notes Tab Panel
        self.treeNotebookNotesTabTreeCtrl = wx.TreeCtrl(self.treeNotebookNotesTab, -1, style = wx.TR_HAS_BUTTONS | wx.TR_EDIT_LABELS )
        # Populate the Tree Control with all the Nodes and Notes it needs!
        self.notesRoot = self.PopulateTreeCtrl(self.treeNotebookNotesTabTreeCtrl)
        # Place the TreeCtrl on the Notes Tab Sizer
        noteTreeCtrlSizer.Add(self.treeNotebookNotesTabTreeCtrl, 1, wx.EXPAND | wx.ALL, 2)
        # Set the Sizer on the Notes Tab
        self.treeNotebookNotesTab.SetSizer(noteTreeCtrlSizer)
        # Fit the controls on the Sizer
        self.treeNotebookNotesTab.Fit()
        # Capture selection event for the Notes Tree Control
        self.treeNotebookNotesTabTreeCtrl.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnTreeItemSelected)
        # Add methods to handle the editing of labels in the Notes Tree
        self.treeNotebookNotesTabTreeCtrl.Bind(wx.EVT_TREE_BEGIN_LABEL_EDIT, self.OnBeginLabelEdit)
        self.treeNotebookNotesTabTreeCtrl.Bind(wx.EVT_TREE_END_LABEL_EDIT, self.OnEndLabelEdit)
        # Add a method to handle right-mouse-down, which should cause a popup menu
        self.treeNotebookNotesTabTreeCtrl.Bind(wx.EVT_RIGHT_DOWN, self.OnTreeCtrlRightDown)

        # Create a Panel for the Notes Search tab of the Notebook
        self.treeNotebookSearchTab = wx.Panel(self.treeNotebook, style = wx.RAISED_BORDER)
        # Add the Notes Search Tab to the Notebook
        self.treeNotebook.AddPage(self.treeNotebookSearchTab, _("Note Search"), False)
        # Create a Sizer for the Notes Search Tab
        noteTreeCtrlSizer2 = wx.BoxSizer(wx.VERTICAL)
        # Create a text header
        prompt = wx.StaticText(self.treeNotebookSearchTab, -1, _("Search Text:"))
        # Add the text to the Note Search Tab sizer
        noteTreeCtrlSizer2.Add(prompt, 0, wx.LEFT | wx.TOP | wx.RIGHT, 2)
        # Add a spacer to the Note Search Tab sizer
        noteTreeCtrlSizer2.AddSpacer((0, 5))
        # Add a Horizontal Sizer for the Search Text and button
        noteSearchSizer = wx.BoxSizer(wx.HORIZONTAL)
        # Create a Search Text Ctrl.  We want this control to handle the Enter key itself
        self.noteSearch = wx.TextCtrl(self.treeNotebookSearchTab, -1, style=wx.TE_PROCESS_ENTER)
        # Capture the Enter key for the Search box.  Enter should trigger a search
        self.noteSearch.Bind(wx.EVT_TEXT_ENTER, self.OnAllNotesSearch)
        # When the noteSearch control loses focus, that should trigger a search too.
        self.noteSearch.Bind(wx.EVT_KILL_FOCUS, self.OnAllNotesSearch)
        # Add the Search Text box to the note search sizer
        noteSearchSizer.Add(self.noteSearch, 1, wx.EXPAND | wx.RIGHT, 5)
        # Get the graphic for the Search button ...
        bmp = wx.ArtProvider_GetBitmap(wx.ART_FIND, wx.ART_TOOLBAR, (16,16))
        # ... and create a button with that graphic
        self.searchButton = wx.BitmapButton(self.treeNotebookSearchTab, -1, bmp, (16, 16))
        # Link the button to the method that will perform the search
        self.searchButton.Bind(wx.EVT_BUTTON, self.OnAllNotesSearch)
        # Add the button to the Note Search Sizer
        noteSearchSizer.Add(self.searchButton, 0)
        # Add the Search Text Siser to the Note Search Tab sizer
        noteTreeCtrlSizer2.Add(noteSearchSizer, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 2)
        # Add a spacer to the Note Search Tab sizer
        noteTreeCtrlSizer2.AddSpacer((0, 10))
        # Create a Tree Control and place it on the Notes Search Tab Panel
        self.treeNotebookSearchTabTreeCtrl = wx.TreeCtrl(self.treeNotebookSearchTab, -1, style = wx.TR_HAS_BUTTONS | wx.TR_EDIT_LABELS )
        # Place the TreeCtrl on the Notes Search Tab Sizer
        noteTreeCtrlSizer2.Add(self.treeNotebookSearchTabTreeCtrl, 1, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 2)
        # Capture selection event for the Notes Tree Control
        self.treeNotebookSearchTabTreeCtrl.Bind(wx.EVT_TREE_SEL_CHANGED, self.OnTreeItemSelected)
        # Add methods to handle the editing of labels in the Notes Tree
        self.treeNotebookSearchTabTreeCtrl.Bind(wx.EVT_TREE_BEGIN_LABEL_EDIT, self.OnBeginLabelEdit)
        self.treeNotebookSearchTabTreeCtrl.Bind(wx.EVT_TREE_END_LABEL_EDIT, self.OnEndLabelEdit)
        # Add a method to handle right-mouse-down, which should cause a popup menu
        self.treeNotebookSearchTabTreeCtrl.Bind(wx.EVT_RIGHT_DOWN, self.OnTreeCtrlRightDown)
        # Set the Sizer on the Notes Search Tab
        self.treeNotebookSearchTab.SetSizer(noteTreeCtrlSizer2)
        # Fit the controls on the Sizer
        self.treeNotebookSearchTab.Fit()

        # Add the Notebook to the Note Tree Panel Sizer
        noteTreeSizer.Add(self.treeNotebook, 1, wx.ALL | wx.EXPAND, 2)
        # Set the Sizer on the Note Tree Panel
        self.treePanel.SetSizer(noteTreeSizer)
        # Fit the controls on the Sizer
        self.treePanel.Fit()

        # Place a Transana Notes Editor Panel in the other half of the Splitter Window
        self.noteEdit = NoteEditor._NotePanel(splitter)
        # Disable the Notes Editor controls initially
        self.noteEdit.EnableControls(False)
        # Set the minimum Pane size for the splitter
        splitter.SetMinimumPaneSize(100)
        # Set the initial split position for the spliter window
        splitter.SplitVertically(self.treePanel, self.noteEdit, 300)
        # Set the Dialog's Main Sizer
        self.SetSizer(mainSizer)
        # Add a Page Changing event to the notebook
        self.treeNotebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGING, self.OnNotebookPageChanging)
        # Add a Page Changed event to the notebook
        self.treeNotebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.OnNotebookPageChanged)
        # Capture the Dialog's Close Event
        wx.EVT_CLOSE(self, self.OnClose)
        # Center the Notes Browser on the screen.
        self.CentreOnScreen()

    def Register(self, ControlObject=None):
        """ Register a Control Object for the Notes Browser to interact with. """
        # Register the Control Object with the Notes Browser Object
        self.ControlObject = ControlObject
        # Register the Notes Browser Object with the Control Object
        if self.ControlObject != None:
            self.ControlObject.Register(NotesBrowser=self)

    def AddNote(self, tree, rootNode, noteNum, noteID, seriesNum, episodeNum, transcriptNum, collectionNum, clipNum, noteTaker):
        # If we have a Series Note ...
        if seriesNum > 0:
            # Load the Series to get the needed data
            tempSeries = Series.Series(seriesNum)
            # The path to the Note is the Series Name
            pathData = tempSeries.id
            # Create the Node Data
            nodeData = DatabaseTreeTab._NodeData('NoteNode', noteNum, seriesNum)
        # If we have an Episode Note ...
        elif episodeNum > 0:
            # Load the Episode and Series to get the needed data
            tempEpisode = Episode.Episode(episodeNum)
            tempSeries = Series.Series(tempEpisode.series_num)
            # The path to the Note is the Series Name and the Episode Name
            pathData = tempSeries.id + ' > ' + tempEpisode.id
            # Create the Node Data
            nodeData = DatabaseTreeTab._NodeData('NoteNode', noteNum, episodeNum)
        # If we have a Transcript Note ...
        elif transcriptNum > 0:
            # Load the Transcript, Episode and Series to get the needed data
            tempTranscript = Transcript.Transcript(transcriptNum)
            tempEpisode = Episode.Episode(tempTranscript.episode_num)
            tempSeries = Series.Series(tempEpisode.series_num)
            # The path to the Note is the Series Name, the Episode Name, and the Transcript Name
            pathData = tempSeries.id + ' > ' + tempEpisode.id + ' > ' + tempTranscript.id
            # Create the Node Data
            nodeData = DatabaseTreeTab._NodeData('NoteNode', noteNum, transcriptNum)
        # If we have a Collection Note ...
        elif collectionNum > 0:
            # Load the Collection to get the needed data
            tempCollection = Collection.Collection(collectionNum)
            # The path to the Note is the Collection's Node String
            pathData = tempCollection.GetNodeString()
            # Create the Node Data
            nodeData = DatabaseTreeTab._NodeData('NoteNode', noteNum, collectionNum)
        # If we have a Clip Note ...
        elif clipNum > 0:
            # Load the Clip and Collection to get the needed data
            tempClip = Clip.Clip(clipNum)
            tempCollection = Collection.Collection(tempClip.collection_num)
            # The path to the Note is the Collection's Node String and the Clip Name
            pathData = tempCollection.GetNodeString() + ' > ' + tempClip.id
            # Create the Node Data
            nodeData = DatabaseTreeTab._NodeData('NoteNode', noteNum, clipNum)
        # Otherwise, we have an undefined note.  (This should never get called)
        else:
            # There is no Root Node
            rootNode = None
            # There is no path to the data
            pathData = ''
            # There is no Node Data
            nodeData = None

        # If we have a Root node ...
        if rootNode != None:
            # ... create the Tree Item for the Note ...
            item = tree.AppendItem(rootNode, noteID)
            # ... add the Item Data ...
            tree.SetPyData(item, nodeData)
            # ... add the path data if defined ...
            if pathData != '':
                item2 = tree.AppendItem(item, pathData)
            # ... and add the note taker information if defined.
            if noteTaker != '':
                item3 = tree.AppendItem(item, noteTaker)

    def PopulateTreeCtrl(self, tree, searchText=None):
        """ Populate the Notes Browser Tree Controls, limiting by SearchText if such a value is passed in """
        # Clear all the nodes from the Tree.  We're starting over.
        tree.DeleteAllItems()
        # Include the Database Name as the Tree Root
        prompt = _('Database: %s')
        # Encode the Database Name if necessary
        if ('unicode' in wx.PlatformInfo):
            prompt = unicode(prompt, 'utf8')
        # Add the Database Name to the prompt
        prompt = prompt % TransanaGlobal.configData.database
        # If searchText is specified, add it to the Tree Root prompt
        if searchText != None:
            prompt += ' ' + unicode(_('Text: %s'), 'utf8') % searchText
        # Add the Tree's root node
        root = tree.AddRoot(prompt)
        tree.SetPyData(root, DatabaseTreeTab._NodeData('RootNode'))
        # Add a Series Node at the first level
        seriesNode = tree.AppendItem(root, _("Series"))
        tree.SetPyData(seriesNode, DatabaseTreeTab._NodeData('SeriesNode'))
        # Add an Episode Node at the first level
        episodeNode = tree.AppendItem(root, _("Episode"))
        tree.SetPyData(episodeNode, DatabaseTreeTab._NodeData('EpisodeNode'))
        # Add a Transcript Node at the first level
        transcriptNode = tree.AppendItem(root, _("Transcript"))
        tree.SetPyData(transcriptNode, DatabaseTreeTab._NodeData('TranscriptNode'))
        # Add a Collection Node at the first level
        collectionNode = tree.AppendItem(root, _("Collection"))
        tree.SetPyData(collectionNode, DatabaseTreeTab._NodeData('CollectionNode'))
        # Add a Clip Node at the first level
        clipNode = tree.AppendItem(root, _("Clip"))
        tree.SetPyData(clipNode, DatabaseTreeTab._NodeData('ClipNode'))

        # Get a list of all Notes from the Database
        notes = DBInterface.list_of_all_notes(searchText=searchText)
        # Iterate through the list of notes
        for note in notes:
            if note['SeriesNum'] > 0:
                rootNode = seriesNode
            elif note['EpisodeNum'] > 0:
                rootNode = episodeNode
            elif note['TranscriptNum'] > 0:
                rootNode = transcriptNode
            elif note['CollectionNum'] > 0:
                rootNode = collectionNode
            elif note['ClipNum'] > 0:
                rootNode = clipNode
            else:
                rootNode = None
            self.AddNote(tree, rootNode, note['NoteNum'], note['NoteID'], note['SeriesNum'], note['EpisodeNum'], note['TranscriptNum'], note['CollectionNum'], note['ClipNum'], note['NoteTaker'])

        # Expand the root node so we see the first level nodes
        tree.Expand(root)
        # Return the root node.
        return root

    def UpdateTreeCtrl(self, action, note, oldName=''):
        """ Update the Tree Control based on an external event, such as multi-user communication """
        # If we're on the Notes tab, not the Note Search tab ...
        if self.treeNotebook.GetCurrentPage() == self.treeNotebookNotesTab:
            # If we're adding a new Note ...
            if action == 'A':
                # Get the Node Data for the Note Object passed in.
                (nodeData, nodeType) = self.GetNodeData(note)
                # Initialize that we want to continue looking for the right node
                contin = True
                # Start at the Root Node of the Notes Tab
                rootNode = self.treeNotebookNotesTabTreeCtrl.GetRootItem()
                # Get the first child from the Root
                (childNode, cookie) = self.treeNotebookNotesTabTreeCtrl.GetFirstChild(rootNode)
                # While we have valid child nodes and have not yet found what we're looking for ...
                while childNode.IsOk() and contin:
                    # Get the Node Data for the child we're currently examining
                    childNodeData = self.treeNotebookNotesTabTreeCtrl.GetPyData(childNode)
                    # See if the Node Type for the Note matches the Node Type for the node we're looking at ...
                    if ((nodeType == 'SeriesNoteNode') and (childNodeData.nodetype == 'SeriesNode')) or \
                       ((nodeType == 'EpisodeNoteNode') and (childNodeData.nodetype == 'EpisodeNode')) or \
                       ((nodeType == 'TranscriptNoteNode') and (childNodeData.nodetype == 'TranscriptNode')) or \
                       ((nodeType == 'CollectionNoteNode') and (childNodeData.nodetype == 'CollectionNode')) or \
                       ((nodeType == 'ClipNoteNode') and (childNodeData.nodetype == 'ClipNode')):
                        # ... if so, add the Note to the child node ...
                        self.AddNote(self.treeNotebookNotesTabTreeCtrl, childNode, note.number, note.id, note.series_num, note.episode_num, note.transcript_num, note.collection_num, note.clip_num, note.author)
                        # ... and indicate we can stop looking
                        contin = False
                    # if we have NOT found what we're looking for ...
                    else:
                        # ... then look at the next child node.
                        (childNode, cookie) = self.treeNotebookNotesTabTreeCtrl.GetNextChild(rootNode, cookie)

            # If we're renaming a Note ...
            elif action == 'R':
                # Get the Node Data for the Note Object passed in.
                (nodeData, nodeType) = self.GetNodeData(note)
                # Initialize that we want to continue looking for the right node
                contin = True
                # Start at the Root Node of the Notes Tab
                rootNode = self.treeNotebookNotesTabTreeCtrl.GetRootItem()
                # Get the first child from the Root
                (childNode, cookie) = self.treeNotebookNotesTabTreeCtrl.GetFirstChild(rootNode)
                # While we have valid child nodes and have not yet found what we're looking for ...
                while childNode.IsOk() and contin:
                    # Get the Node Data for the child we're currently examining
                    childNodeData = self.treeNotebookNotesTabTreeCtrl.GetPyData(childNode)
                    # See if the Node Type for the Note matches the Node Type for the node we're looking at ...
                    if ((nodeType == 'SeriesNoteNode') and (childNodeData.nodetype == 'SeriesNode')) or \
                       ((nodeType == 'EpisodeNoteNode') and (childNodeData.nodetype == 'EpisodeNode')) or \
                       ((nodeType == 'TranscriptNoteNode') and (childNodeData.nodetype == 'TranscriptNode')) or \
                       ((nodeType == 'CollectionNoteNode') and (childNodeData.nodetype == 'CollectionNode')) or \
                       ((nodeType == 'ClipNoteNode') and (childNodeData.nodetype == 'ClipNode')):
                        # ... if so, start looking at the children of this node ...
                        (noteNode, cookie2) = self.treeNotebookNotesTabTreeCtrl.GetFirstChild(childNode)
                        # While we have valid grandchild nodes and have not yet found what we're looking for ...
                        while noteNode.IsOk() and contin:
                            # ... see if the node text matches the Note's OLD name (a parameter used only for renaming!)
                            if self.treeNotebookNotesTabTreeCtrl.GetItemText(noteNode) == oldName:
                                # If it matches, update the Node's text to the Note's NEW name ...
                                self.treeNotebookNotesTabTreeCtrl.SetItemText(noteNode, note.id)
                                # ... and signal that we're done.
                                contin = False
                            # If the text is NOT a match ...
                            else:
                                # ... continue on to the next grandchild node
                                (noteNode, cookie2) = self.treeNotebookNotesTabTreeCtrl.GetNextChild(childNode, cookie2)
                    # if we have NOT found what we're looking for ...
                    else:
                        # ... then look at the next child node.
                        (childNode, cookie) = self.treeNotebookNotesTabTreeCtrl.GetNextChild(rootNode, cookie)

            # If we're deleting a Note ...
            elif action == 'D':
                # Initialize that we want to continue looking for the right node
                contin = True
                # Start at the Root Node of the Notes Tab
                rootNode = self.treeNotebookNotesTabTreeCtrl.GetRootItem()
                # Get the first child from the Root
                (childNode, cookie) = self.treeNotebookNotesTabTreeCtrl.GetFirstChild(rootNode)
                # While we have valid child nodes and have not yet found what we're looking for ...
                while childNode.IsOk() and contin:
                    # See if the node's text matches the first part of the note tuple (not a note object) passed.
                    # (Translate the object type and convert it to Unicode HERE, not before sending!)
                    if self.treeNotebookNotesTabTreeCtrl.GetItemText(childNode) == unicode(_(note[0]), 'utf8'):
                        # If so, get the first grandchild (note) node
                        (noteNode, cookie2) = self.treeNotebookNotesTabTreeCtrl.GetFirstChild(childNode)
                        # While we have valid grandchild nodes and have not yet found what we're looking for ...
                        while noteNode.IsOk() and contin:
                            # ... see if the node text matches the Note's ID (passed as the second part of the "note" tuple!)
                            if self.treeNotebookNotesTabTreeCtrl.GetItemText(noteNode) == note[1]:
                                # If so, remove that node from the tree ...
                                self.treeNotebookNotesTabTreeCtrl.Delete(noteNode)
                                # ... and signal that we're done.
                                contin = False
                            # If the text is NOT a match ...
                            else:
                                # ... continue on to the next grandchild node
                                (noteNode, cookie2) = self.treeNotebookNotesTabTreeCtrl.GetNextChild(childNode, cookie2)
                    # If the text is NOT a match ...
                    else:
                        # ... then look at the next child node.
                        (childNode, cookie) = self.treeNotebookNotesTabTreeCtrl.GetNextChild(rootNode, cookie)

            # This should NEVER get triggered!
            else:
                print 'NotesBrowser.UpdateTreeCtrl():  Unknown action "%s"' % action
                print note
                print

    def GetNodeData(self, note, idChanged=False):
        """ Return the Node Data for the Note object passed in.  idChanged indicates whether to use the note ID
            or its original value, which could be different. """
        # Initialize the return values to blanks
        nodeData = ''
        nodeType = ''
        # If it's possible the Note ID has been changed ...
        if idChanged:
            # use the saved Original Note ID
            noteIDToUse = self.originalNoteID
        # If it's NOT possible the note ID has been changed ...
        else:
            # ... then we can just use the Note ID.
            # (When working with a note locked by another user, originalNoteID may be BLANK!)
            noteIDToUse = note.id
        # Determine what type of note we have and set the appropriate data.
        # If we have a Series note ...
        if note.series_num != 0:
            # ... load the Series data ...
            tempSeries = Series.Series(note.series_num)
            # ... set up the Tree Node data ...
            nodeData = ("Series", tempSeries.id, noteIDToUse)
            # ... and signal that it's a Series Note.
            nodeType = 'SeriesNoteNode'
        # if we have an Episode note ...
        elif note.episode_num != 0:
            # ... load the Series and Episode data ...
            tempEpisode = Episode.Episode(note.episode_num)
            tempSeries = Series.Series(tempEpisode.series_num)
            # ... set up the Tree Node data ...
            nodeData = ("Series", tempSeries.id, tempEpisode.id, noteIDToUse)
            # ... and signal that it's an Episode Note.
            nodeType = 'EpisodeNoteNode'
        # if we have a Transcript note ...
        elif note.transcript_num != 0:
            # ... load the Series, Episode, and Transcript data ...
            tempTranscript = Transcript.Transcript(note.transcript_num)
            tempEpisode = Episode.Episode(tempTranscript.episode_num)
            tempSeries = Series.Series(tempEpisode.series_num)
            # ... set up the Tree Node data ...
            nodeData = ("Series", tempSeries.id, tempEpisode.id, tempTranscript.id, noteIDToUse)
            # ... and signal that it's a Transcript Note.
            nodeType = 'TranscriptNoteNode'
        # if we have a Collection note ...
        elif note.collection_num != 0:
            # ... load the Collection data ...
            tempCollection = Collection.Collection(note.collection_num)
            # ... set up the Tree Node data ...
            nodeData = ("Collections",) + tempCollection.GetNodeData() + (noteIDToUse,)
            # ... and signal that it's a Collection Note.
            nodeType = 'CollectionNoteNode'
        # if we have a Clip note ...
        elif note.clip_num != 0:
            # ... load the Collection and Clip data ...
            tempClip = Clip.Clip(note.clip_num)
            tempCollection = Collection.Collection(tempClip.collection_num)
            # ... set up the Tree Node data ...
            nodeData = ("Collections",) + tempCollection.GetNodeData() + (tempClip.id, noteIDToUse,)
            # ... and signal that it's a Clip Note.
            nodeType = 'ClipNoteNode'
        return (nodeData, nodeType)

    def SaveNoteAndClear(self):
        # Let's check to see if we have a note that is already open.
        if (self.activeNote != None):
            # Start exception handling in case there's a SaveError
            try:
                # DON'T CHECK if the note's changed here.  The ID could be changed!
                # ... get the text from the note edit control ...
                self.activeNote.text = self.noteEdit.get_text()
                # ... and save the active note.
                self.activeNote.db_save()
                # Unlock the record for the active note
                self.activeNote.unlock_record()
                # Clear the active node
                self.activeNote = None
                self.originalNoteID = ''
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()

        # Reset the Note Editor to blank
        self.noteEdit.set_text('')
        # Disable the Note Editor controls
        self.noteEdit.EnableControls(False)

    def OnNotebookPageChanging(self, event):
        """ This method is triggered when the user starts to change Notebook pages """
        # When we start changing tabs, we need to save and clear the active note
        self.SaveNoteAndClear()

    def OnNotebookPageChanged(self, event):
        """ This method is triggered when the user has changed Notebook pages """
        # Allow the default method to process as needed (so the Notebook page changes)
        event.Skip()
        # Determine which page we've changed to and update the Tree control.  (If we renamed something, it'll be
        # out of date on the tab.
        if self.treeNotebook.GetPageText(event.GetSelection()) == unicode(_("Notes"), 'utf8'):
            self.notesRoot = self.PopulateTreeCtrl(self.treeNotebookNotesTabTreeCtrl)
        elif self.treeNotebook.GetPageText(event.GetSelection()) == unicode(_("Note Search"), 'utf8'):
            self.searchRoot = self.PopulateTreeCtrl(self.treeNotebookSearchTabTreeCtrl, self.noteSearch.GetValue())


    def OnTreeItemSelected(self, event):
        """ Process selection of a tree node """
        # Determine which tree is the source of the selection, the Notes tree or the Note Search tree
        if event.GetId() == self.treeNotebookNotesTabTreeCtrl.GetId():
            tree = self.treeNotebookNotesTabTreeCtrl
        elif event.GetId() == self.treeNotebookSearchTabTreeCtrl.GetId():
            tree = self.treeNotebookSearchTabTreeCtrl
        else:
            tree = None
        # If we know the source of the selection ...
        if tree != None:
            # Save the Active Note, if there is one
            self.SaveNoteAndClear()
            # Get the selected item
            sel_item = tree.GetSelection()
            # Get the selected item's data
            sel_item_data = tree.GetPyData(sel_item)
            # If we have selected a node that HAS data, and it's a NOTE node ...
            if (sel_item_data != None) and (sel_item_data.nodetype == 'NoteNode'):
                # We need exception handling here 
                try:
                    # Create a Note object and populated it for the selected note
                    self.activeNote = Note.Note(sel_item_data.recNum)
                    # Try to lock the note
                    self.activeNote.lock_record()
                    # Remember the note's ID so we can tell later if it's changed.
                    self.originalNoteID = self.activeNote.id
                    # If successful, enable the Note Editor controls ...
                    self.noteEdit.EnableControls(True)
                    # ... and populate the Note Editor with the note's text
                    self.noteEdit.set_text(self.activeNote.text)
                    # Make sure the selected object is expanded
                    tree.Expand(sel_item)
                    # If we're on the Note Search tab ...
                    if (tree == self.treeNotebookSearchTabTreeCtrl):
                        # ... then take the search text and stick it in the Note Search Text box too!
                        self.noteEdit.SetSearchText(self.noteSearch.GetValue())
                # Trap Record Lock exceptions
                except TransanaExceptions.RecordLockedError, err:
                    # If the note is locked, we need to inform the user of that fact.  First create a prompt
                    # and encode it as needed.
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Note "%s" is locked by %s\n'), 'utf8')
                    else:
                        prompt = _('Note "%s" is locked by %s\n')
                    # Populate the read-only Note Editor with the lock prompt and the note text.
                    self.noteEdit.set_text(prompt % (self.activeNote.id, self.activeNote.record_lock) +
                                           '\n\n%s' % self.activeNote.text)
                    # Since the record is locked and the control read-only, we can immediately forget that we
                    # have an active note.
                    self.activeNote = None
                # Trap Record Not Found exception, which is triggered if a different user deletes a Note's Parent, grandparent, etc.
                # while this user has the Notes Browser open, then this user selects the deleted Note.  (Deleting the note itself
                # causes it to be removed form the Notes Browser, but not deleting its ancestors.)
                except TransanaExceptions.RecordNotFoundError, err:
                    # We can immediately forget that we have an active note, since the note doesn't exist!
                    self.activeNote = None
                    # Display an error message to the user
                    errordlg = Dialogs.ErrorDialog(None, err.args)
                    errordlg.ShowModal()
                    errordlg.Destroy()

    def OnBeginLabelEdit(self, event):
        """ Handle the start of editing Tree Labels """
        # Determine which tree is the source of the selection, the Notes tree or the Note Search tree
        if event.GetId() == self.treeNotebookNotesTabTreeCtrl.GetId():
            tree = self.treeNotebookNotesTabTreeCtrl
        elif event.GetId() == self.treeNotebookSearchTabTreeCtrl.GetId():
            tree = self.treeNotebookSearchTabTreeCtrl
        else:
            tree = None
        # If we know the source of the selection ...
        if tree != None:
            # Get the current selected item.
            sel_item = tree.GetSelection()
            # Get the item data for the current selection
            sel_item_data = tree.GetPyData(sel_item)
            # If the item data is None or not a Note Node or the note is locked by someone else ...
            if (sel_item_data == None) or (sel_item_data.nodetype != 'NoteNode') or (self.activeNote == None) or \
               (not self.activeNote.isLocked):
                # ... we need to cancel the Label Edit
                event.Veto()
        # If we don't have a known tree ...
        else:
            # ... we need to cancel the Label Edit
            event.Veto()
        

    def OnEndLabelEdit(self, event):
        """ Handle the end of editing Tree Labels """
        # Determine which tree is the source of the selection, the Notes tree or the Note Search tree
        if event.GetId() == self.treeNotebookNotesTabTreeCtrl.GetId():
            tree = self.treeNotebookNotesTabTreeCtrl
        elif event.GetId() == self.treeNotebookSearchTabTreeCtrl.GetId():
            tree = self.treeNotebookSearchTabTreeCtrl
        else:
            tree = None
        # If we know the source of the selection ...
        if tree != None:
            # Get the current selected item.
            sel_item = tree.GetSelection()
            # Get the item data for the current selection
            sel_item_data = tree.GetPyData(sel_item)
            # If there is an active Note and we have it locked and the user didn't cancel the edit...
            if (self.activeNote != None) and self.activeNote.isLocked and not event.IsEditCancelled():
                # ... update the note ID to match the changed Label
                self.activeNote.id = event.GetLabel().strip()
                # Start exception handling
                try:
                    # We need to save here, so we can still veto the change if there's a save error!
                    self.activeNote.db_save()
                    # If we've changed the Note ID ...
                    if self.activeNote.id != self.originalNoteID:
                        # ... and we have a known ControlObject ...
                        if self.ControlObject != None:
                            # ... determine what type of note we have and set the appropriate data.  It's LIKELY
                            # that the node ID has changed here.
                            (nodeData, nodeType) = self.GetNodeData(self.activeNote, idChanged=True)
                            # As long as we have good data ...
                            if nodeData != '':
                                # inform the Database tree of the Note ID change.  (We need to translate the Root node.)
                                self.ControlObject.DataWindow.DBTab.tree.rename_Node((unicode(_(nodeData[0]), 'utf8'),) + nodeData[1:], nodeType, self.activeNote.id)
                                # If we have a Chat Window (and thus are in MU) ...
                                if TransanaGlobal.chatWindow != None:
                                    # Start building a Rename message with the Rename header and the node type.
                                    msg = "RN %s >|< " % nodeType
                                    # Add all the nodes, which we already have assembled ...
                                    for node in nodeData:
                                        msg += node + ' >|< '
                                    # ... and finish the message with the new name.
                                    msg += self.activeNote.id
                                    # Now send the message via the Chat Window.
                                    TransanaGlobal.chatWindow.SendMessage(msg)
                # Duplicate Note IDs or other problems can cause SaveError exceptions
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Veto the Label Edit, since it was rejected
                    event.Veto()
                    # Restore the ActiveNote's ID to the original value
                    self.activeNote.id = self.originalNoteID
            # If there's a problem ...
            else:
                # ... we need to cancel the Label Edit
                event.Veto()
        # if we don't have a known tree ...
        else:
            # ... we need to cancel the Label Edit.  (We should never get here!)
            event.Veto()

    def OnAllNotesSearch(self, event):
        """ "Search All Notes" has been activated.  We search for a string across all notes. """
        # If the search text has changed ...
        if self.searchText != self.noteSearch.GetValue():
            # ... update the variable that remembers our search term
            self.searchText = self.noteSearch.GetValue()
            # If it's empty ...
            if self.searchText == '':
                # ... we need to pass "None" rather than an empty string so the query works right.
                self.searchText = None
            # Populate the Tree Control with all the Nodes and Notes it needs!
            self.searchRoot = self.PopulateTreeCtrl(self.treeNotebookSearchTabTreeCtrl, self.searchText)

    def OnTreeCtrlRightDown(self, event):
        """ Implement the right-click menus for the Tree Control """
        # Determine which tree is the source of the selection, the Notes tree or the Note Search tree
        if event.GetId() == self.treeNotebookNotesTabTreeCtrl.GetId():
            self.activeTree = self.treeNotebookNotesTabTreeCtrl
        elif event.GetId() == self.treeNotebookSearchTabTreeCtrl.GetId():
            self.activeTree = self.treeNotebookSearchTabTreeCtrl
        else:
            self.activeTree = None
        # If we know the source of the selection ...
        if self.activeTree != None:
            # Items in the tree are not automatically selected with a right click.
            # We must select the item that is initially clicked manually!!
            # We do this by looking at the screen point clicked and applying the tree's
            # HitTest method to determine the current item, then actually selecting the item

            # Get the Mouse Position on the Screen in a more generic way to avoid the problem above
            (windowx, windowy) = wx.GetMousePosition()
            # Translate the Mouse's Screen Position to the Mouse's Control Position
            pt = self.activeTree.ScreenToClientXY(windowx, windowy)
            # use HitTest to determine the tree item as the screen point indicated.
            sel_item, flags = self.activeTree.HitTest(pt)
            # Select the appropriate item in the TreeCtrl
            self.activeTree.SelectItem(sel_item)
            # Get the item data for the current selection
            sel_item_data = self.activeTree.GetPyData(sel_item)
            # Determine what type of node we've selected so we can create the correct menu.
            # First, eliminate nodes that don't have a node type, as they don't ge a menu.
            if sel_item_data != None:
                # Initialize the menu to None
                menu = None
                # If we have the Root Node ...
                if sel_item_data.nodetype == 'RootNode':
                    # ... create the menu object
                    menu = wx.Menu()
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for the Notes Report
                    menu.Append(id, _("Notes Report"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuReport)
                # If we have the Series Root Node ...
                elif sel_item_data.nodetype == 'SeriesNode':
                    # ... create the menu object
                    menu = wx.Menu()
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for the Series Notes Report
                    menu.Append(id, _("Series Notes Report"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuReport)
                # If we have the Episode Root Node ...
                elif sel_item_data.nodetype == 'EpisodeNode':
                    # ... create the menu object
                    menu = wx.Menu()
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for the Episode Notes Report
                    menu.Append(id, _("Episode Notes Report"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuReport)
                # If we have the Transcript Root Node ...
                elif sel_item_data.nodetype == 'TranscriptNode':
                    # ... create the menu object
                    menu = wx.Menu()
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for the Transcript Notes Report
                    menu.Append(id, _("Transcript Notes Report"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuReport)
                # If we have the Collection Root Node ...
                elif sel_item_data.nodetype == 'CollectionNode':
                    # ... create the menu object
                    menu = wx.Menu()
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for the Collection Notes Report
                    menu.Append(id, _("Collection Notes Report"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuReport)
                # If we have the Clip Root Node ...
                elif sel_item_data.nodetype == 'ClipNode':
                    # ... create the menu object
                    menu = wx.Menu()
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for the Clip Notes Report
                    menu.Append(id, _("Clip Notes Report"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuReport)
                # If we have a Note Node ...
                elif sel_item_data.nodetype == 'NoteNode':
                    # ... create the menu object
                    menu = wx.Menu()
                    # See if we're pointing to a Clip Note by looking at the parent's label.  If so ...
                    if (self.activeTree.GetItemText(self.activeTree.GetItemParent(sel_item)) == unicode(_('Clip'), 'utf8')):
                        # Create a new Id for a menu item
                        id = wx.NewId()
                        # Add a menu item for loading a Clip Note's Clip
                        menu.Append(id, _("Load the Associated Clip"))
                        # Link the appropriate method to the menu item
                        wx.EVT_MENU(self, id, self.OnMenuLoadClip)
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for Locating the Notes in the DB Tree
                    menu.Append(id, _("Locate Note"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuLocate)
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for renaming the Note
                    menu.Append(id, _("Rename"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuRename)
                    # If there is no activeNote (i.e. the note is locked by someone else) ...
                    if (self.activeNote == None):
                        # ... then Note Properties should not be available as a menu option
                        menu.Enable(id, False)
                    # Create a new Id for a menu item
                    id = wx.NewId()
                    # Add a menu item for Note Properties
                    menu.Append(id, _("Note Properties"))
                    # Link the appropriate method to the menu item
                    wx.EVT_MENU(self, id, self.OnMenuProperties)
                    # If there is no activeNote (i.e. the note is locked by someone else) ...
                    if (self.activeNote == None):
                        # ... then Note Properties should not be available as a menu option
                        menu.Enable(id, False)
                # If a Menu Object was created ...
                if menu != None:
                    # ... show the menu!  (We have to adjust the position of the popup for the position of the tree control.)
                    self.PopupMenu(menu, (event.GetPosition()[0], event.GetPosition()[1] + self.activeTree.GetPositionTuple()[1] + 15))

    def OnMenuReport(self, event):
        """ Implement the Notes Report function from the right-click menus """
        # Determine if this is called from the Note Search tree.
        if (self.activeTree.GetId() == self.treeNotebookSearchTabTreeCtrl.GetId()):
            # If so, we need to sent the search text.
            searchText = self.searchText
        # If not ...
        else:
            # ... set the search text to None
            searchText = None
        # Call the Notes Report, passing the node type and the search text
        ReportGeneratorForNotes.ReportGenerator(title=_("Transana Notes Report"),
                                                reportType=self.activeTree.GetPyData(self.activeTree.GetSelection()).nodetype,
                                                searchText=searchText)
        
    def OnMenuLoadClip(self, event):
        """ Implement the Load Clip function from the right-click menu """
        # See if the current note is defined (which it isn't if locked by someone else!)
        if self.activeNote != None:
            # Get the Note's Node Data ...
            (nodeData, nodeType) = self.GetNodeData(self.activeNote)
            # ... and remember the clip's number
            clipNum = self.activeNote.clip_num
        # If the note is locked, we can still load the Clip.  But we don't have the activeNote to get data from.
        else:
            # Get the selected item
            sel_item = self.activeTree.GetSelection()
            # Get the selected item's data
            sel_item_data = self.activeTree.GetPyData(sel_item)
            # Since the note is locked, there is no activeNote.  So we'll just get the note data temporarily.
            tempNote = Note.Note(sel_item_data.recNum)
            # Get the Note's Node Data ...
            (nodeData, nodeType) = self.GetNodeData(tempNote)
            # ... and remember the clip's number
            clipNum = tempNote.clip_num
        # We want the Clip, not the Clip Note!  We need to drop the note from the list and change the node type!
        nodeData = nodeData[:-1]
        # Set the node type to Clip, not Clip Note
        nodeType = 'ClipNode'
        # Tell the Database Tree to display that note.  (We need to translate the Root node.)
        self.ControlObject.DataWindow.DBTab.tree.select_Node((unicode(_(nodeData[0]), 'utf8'),) + nodeData[1:], nodeType)
        # Load the Clip in Transana's main interface
        self.ControlObject.LoadClipByNumber(clipNum)
        # Close the Notes Browser so user and view the clip
        self.Close()
        
    def OnMenuLocate(self, event):
        """ Implement the Locate Note function from the right-click menu """
        # See if the current note is defined (which it isn't if locked by someone else!)
        if self.activeNote != None:
            # Get the Note's Node Data
            (nodeData, nodeType) = self.GetNodeData(self.activeNote)
        # If the note is locked, we can still locate the Clip Note.  But we don't have the activeNote to get data from.
        else:
            # Get the selected item
            sel_item = self.activeTree.GetSelection()
            # Get the selected item's data
            sel_item_data = self.activeTree.GetPyData(sel_item)
            # Since the note is locked, there is no activeNote.  So we'll just get the note data temporarily.
            tempNote = Note.Note(sel_item_data.recNum)
            # Get the Note's Node Data
            (nodeData, nodeType) = self.GetNodeData(tempNote)
        # Tell the Database Tree to display that note.  (We need to translate the Root node.)
        self.ControlObject.DataWindow.DBTab.tree.select_Node((unicode(_(nodeData[0]), 'utf8'),) + nodeData[1:], nodeType)
        
    def OnMenuRename(self, event):
        """ Implement the Rename function from the right-click menu """
        # Get the current selection from the active Tree Ctrl
        sel_item = self.activeTree.GetSelection()
        # Tell that Tree Ctrl to initiate Editing the Label!
        self.activeTree.EditLabel(sel_item)

    def OnMenuProperties(self, event):
        """ Implement the Note Properties function from the right-click menu """
        # You can only do this if the current note is defined (which it isn't if locked by someone else!)
        if self.activeNote != None:
            # Get the current selection from the active Tree Ctrl
            sel_item = self.activeTree.GetSelection()
            # Remember the original value of author
            originalNoteAuthor = self.activeNote.author
            # Create the Note Properties Dialog Box to edit the Note Properties
            dlg = NotePropertiesForm.EditNoteDialog(self, -1, self.activeNote)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # While the "continue" flag is True ...
            while contin:
                # Initiate exception handling
                try:
                    # Display the Note Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # If the user says OK, try to save the Note Data
                        self.activeNote.db_save()
                        # See if the Note ID has been changed.  If it has ...
                        if self.activeNote.id != self.originalNoteID:
                            # ... update the Notes Browser tree ...
                            self.activeTree.SetItemText(sel_item, self.activeNote.id)
                            # ... get the note's Node Data (The Node ID could well have been changed) ...
                            (nodeData, nodeType) = self.GetNodeData(self.activeNote, idChanged=True)
                            # ... and inform the Database tree of the Note ID change.  (We need to translate the Root node.)
                            self.ControlObject.DataWindow.DBTab.tree.rename_Node((unicode(_(nodeData[0]),'utf8'),) + nodeData[1:], nodeType, self.activeNote.id)
                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # If we have a Chat Window (and thus are in MU) ...
                                if TransanaGlobal.chatWindow != None:
                                    # Start building a Rename message with the node type.
                                    msg = "%s >|< " % nodeType
                                    # Add all the nodes, which we already have assembled ...
                                    for node in nodeData:
                                        msg += node + ' >|< '
                                    # ... and finish the message with the new name.
                                    msg += self.activeNote.id
                                    if DEBUG:
                                        print 'Message to send = "RN %s"' % msg
                                    # Send the Rename Node message
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)
                        # See if the Author has changed.  If so, we need to reflect it in the Notes Tree.
                        if self.activeNote.author != originalNoteAuthor:
                            # Initialize that the change has not yet been made.
                            changeMade = False
                            # We need to locate the node that has the previous Author information, if it exists, which it might not.
                            # Start with the first child of the currently selected node.
                            (childNode, cookie) = self.activeTree.GetFirstChild(sel_item)
                            # Keep going as long as we have a valid child (not at the end of the list) and the update hasn't already
                            # been made.
                            while childNode.IsOk() and not changeMade:
                                # If we've found the node that matches the OLD author ...
                                if self.activeTree.GetItemText(childNode) == originalNoteAuthor:
                                    # ... if author hasn't been removed ...
                                    if self.activeNote.author.strip() != '':
                                        # ... update the author node.
                                        self.activeTree.SetItemText(childNode, self.activeNote.author)
                                    # ... if the author HAS been removed ...
                                    else:
                                        # ... then we need to delete the tree node where it used to be.
                                        self.activeTree.Delete(childNode)
                                    # Flag that the change has been made, so we can stop the loop.
                                    changeMade = True
                                # If we didn't match the old author ...
                                else:
                                    # ... look for the next node.  If this hits the end of the list, the loop will not pass the IsOk() test.
                                    (childNode, cookie) = self.activeTree.GetNextChild(sel_item, cookie)
                            # If the change wasn't made and we have an author ...
                            if not changeMade and (self.activeNote.author.strip() != ''):
                                # ... there must not have been an author node, so add one to the selected item.
                                self.activeTree.AppendItem(sel_item, self.activeNote.author)
                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except TransanaExceptions.SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                # Handle other exceptions
                except:
                    if DEBUG:
                        import traceback
                        traceback.print_exc(file=sys.stdout)
                    # Display the Exception Message, allow "continue" flag to remain true
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_("Exception %s: %s"), 'utf8')
                    else:
                        prompt = _("Exception %s: %s")
                    errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                    errordlg.ShowModal()
                    errordlg.Destroy()

    def OnHelp(self, event):
        """ Implement the Help function, as required by the Note Editor """
        # If a Help Window and the Help Context are defined ...
        if (TransanaGlobal.menuWindow != None):
            # ... call Help!
            TransanaGlobal.menuWindow.ControlObject.Help("NotesBrowser")


    def OnClose(self, event):
        """ Implement the Window Close function, as required by the Note Editor """
        # Save the active note, if there is one
        self.SaveNoteAndClear()
        # Remove the reference to the Notes Browser from the Control Object
        if self.ControlObject != None:
            self.ControlObject.Register(NotesBrowser=None)
            
        # self.Close() leads to recursion!  But self.Destroy() seems to cause a seg fault on the Mac.
        # Let's just stop showing the window, and let the calling routine destroy the dialog.
        self.Show(False)

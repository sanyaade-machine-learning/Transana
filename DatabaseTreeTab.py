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

"""This module implements the DatabaseTreeTab class for the Data Display Objects."""

__author__ = 'David Woods <dwoods@wcer.wisc.edu>, Nathaniel Case'

DEBUG = False
if DEBUG:
    print "DatabaseTreeTab DEBUG is ON!!"

import wx
import TransanaConstants
import TransanaGlobal
import TransanaImages
import Library
import LibraryPropertiesForm
import Document
import Episode
import EpisodePropertiesForm
import Transcript
import TranscriptPropertiesForm
import Collection
import CollectionPropertiesForm
import Quote
import Clip
import Snapshot
import ClipPropertiesForm
if TransanaConstants.proVersion:
    import DocumentPropertiesForm
    import QuotePropertiesForm
    import SnapshotPropertiesForm
import Note
import NotePropertiesForm
import NotesBrowser
import KeywordObject as Keyword                      
import KeywordPropertiesForm
import ClipKeywordObject
import BatchFileProcessor
import ProcessSearch
from TransanaExceptions import *
import KWManager
import exceptions
import DBInterface
import Dialogs
import NoteEditor
# import Python's datetime module
import datetime
import os
import sys
import string
import time
import LibraryMap
import KeywordMapClass
import ReportGenerator
import KeywordSummaryReport
import AnalyticDataExport
import PlayAllClips
import DragAndDropObjects           # Implements Drag and Drop logic and objects
import cPickle                      # Used in Drag and Drop
import Misc                         # Transana's Miscellaneous functions
import PropagateChanges             # Transana's Change Propagation routines
import MediaConvert

class DatabaseTreeTab(wx.Panel):
    """This class defines the object for the "Database" tab of the Data
    window."""
    def __init__(self, parent):
        """Initialize a DatabaseTreeTab object."""
        self.parent = parent
        psize = parent.GetSizeTuple()
        width = psize[0] - 13 
        height = psize[1] - 45

        self.ControlObject = None            # The ControlObject handles all inter-object communication, initialized to None

        # Use WANTS_CHARS style so the panel doesn't eat the Enter key.
        wx.Panel.__init__(self, parent, -1, style=wx.WANTS_CHARS, size=(width, height), name='DatabaseTreeTabPanel')

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        tID = wx.NewId()

        self.tree = _DBTreeCtrl(self, tID, wx.DefaultPosition, self.GetSizeTuple(), wx.TR_HAS_BUTTONS | wx.TR_EDIT_LABELS | wx.TR_MULTIPLE)

        mainSizer.Add(self.tree, 1, wx.EXPAND)
        self.tree.UnselectAll()
        self.tree.SelectItem(self.tree.GetRootItem())

        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        

    def Register(self, ControlObject=None):
        """ Register a ControlObject  for the DataaseTreeTab to interact with. """
        self.ControlObject=ControlObject


    def add_series(self):
        """User interface for adding a new Library."""
        # Create the Library Properties Dialog Box to Add a Library
        dlg = LibraryPropertiesForm.AddLibraryDialog(self, -1)
        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Use "try", as exceptions could occur
            try:
                # Display the Library Properties Dialog Box and get the data from the user
                library = dlg.get_input()
                # If the user pressed OK ...
                if library != None:
                    # Try to save the data from the form
                    self.save_series(library)
                    nodeData = (_('Libraries'), library.id)
                    self.tree.add_Node('LibraryNode', nodeData, library.number, 0)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        if DEBUG:
                            print 'Message to send = "AS %s"' % nodeData[-1]
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("AS %s" % nodeData[-1])

                    # If we do all this, we don't need to continue any more.
                    contin = False
                # If the user pressed Cancel ...
                else:
                    # ... then we don't need to continue any more.
                    contin = False
            # Handle "SaveError" exception
            except SaveError:
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
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
                                    
        
    def edit_series(self, library):
        """User interface for editing a Library."""
        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            library.lock_record()
        # Handle the exception if the record is locked
        except RecordLockedError, e:
            self.handle_locked_record(e, _('Libraries'), library.id)
        # If the record is not locked, keep going.
        else:
            # Create the Library Properties Dialog Box to edit the Library Properties
            dlg = LibraryPropertiesForm.EditLibraryDialog(self, -1, library)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the current tree node
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Library Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # Save the Library
                        self.save_series(library)
                        # remember the node's original name
                        originalName = self.tree.GetItemText(sel)
                        # See if the Library ID has been changed.  If it has, update the tree.
                        if library.id != originalName:
                            # Change the name of the current tree node
                            self.tree.SetItemText(sel, library.id)
                            # If we're in Multi-user mode, we need to send a message about the change.
                            if not TransanaConstants.singleUserVersion:
                                # Build the message.  The first element is the node type, the second element is
                                # the UNTRANSLATED Root Node name in order to avoid problems in mixed-language environents.
                                # the last element is the new name.
                                msg = "%s >|< %s >|< %s >|< %s" % ('LibraryNode', 'Libraries', originalName, library.id)
                                if DEBUG:
                                    if 'unicode' in wx.PlatformInfo:
                                        print 'Message to send = "RN %s"' % msg.encode('utf_8')
                                    else:
                                        print 'Message to send = "RN %s"' % msg
                                # Send the message.
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
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
                        prompt = unicode(_("Exception %s : %s"), 'utf8')
                    else:
                        prompt = _("Exception %s : %s")
                    errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                    errordlg.ShowModal()
                    errordlg.Destroy()
            # Unlock the record regardless of what happens
            library.unlock_record()

    def save_series(self, library):
        """Save/Update the Library object."""
        # FIXME: Graceful exception handling
        if library != None:
            library.db_save()
 
    def add_document(self, library_name):
        """User interface for adding a new document."""
        # Load the Library which contains the Document
        library = Library.Library(library_name)
        try:
            # Lock the Library, to prevent it from being deleted out from under the add.
            library.lock_record()
            libraryLocked = True
        # Handle the exception if the record is already locked by someone else
        except RecordLockedError, s:
            # If we can't get a lock on the Library, it's really not that big a deal.  We only try to get it
            # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
            # still be able to add Documents even if someone else is editing the Library properties.
            libraryLocked = False
        # Create the Document Properties Dialog Box to Add a Document
        dlg = DocumentPropertiesForm.AddDocumentDialog(self, -1, library)
        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Use "try", as exceptions could occur
            try:
                # Display the Document Properties Dialog Box and get the data from the user
                document = dlg.get_input()
                # If the user pressed OK ...
                if document != None:
                    # Try to save the data from the form
                    self.save_document(document)
                    nodeData = (_('Libraries'), library.id, document.id)
                    self.tree.add_Node('DocumentNode', nodeData, document.number, library.number)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        if DEBUG:
                            print 'Message to send = "AD %s >|< %s"' % (nodeData[-2], nodeData[-1])
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("AD %s >|< %s" % (nodeData[-2], nodeData[-1]))

                    # If we do all this, we don't need to continue any more.
                    contin = False
                    # Unlock the Library, if we locked it.
                    if libraryLocked:
                        library.unlock_record()
                    # Return the Document Name
                    return document.id
                # If the user pressed Cancel ...
                else:
                    # ... then we don't need to continue any more.
                    contin = False
                    # Unlock the Library, if we locked it.
                    if libraryLocked:
                        library.unlock_record()
                    # If no Document is created, indicate this.
                    return None
            # Handle "SaveError" exception
            except SaveError:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()

                # Refresh the Keyword List, if it's a changed Keyword error
                dlg.refresh_keywords()
                # Highlight the first non-existent keyword in the Keywords control
                dlg.highlight_bad_keyword()

            # Handle other exceptions
            except:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    
                # Display the Exception Message, allow "continue" flag to remain true
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()

    def edit_document(self, document):
        """User interface for editing a document."""

        # If the user tries to edit a currently-loaded Document, close it first.
        self.ControlObject.CloseOpenTranscriptWindowObject(Document.Document, document.number)
        # This could have let to a change in the document object!  We better re-load it just to be safe.
        document.db_load_by_num(document.number)
        
        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            document.lock_record()
        # Handle the exception if the record is already locked by someone else
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Document"), document.id)
        # If the record is not locked, keep going.
        else:
            # Create the Document Properties Dialog Box to edit the Document Properties
            dlg = DocumentPropertiesForm.EditDocumentDialog(self, -1, document)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection (Before calling up the Properties Form, which can interfere with this!)
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Document Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # Check to see if there are keywords to be propagated
                        self.ControlObject.PropagateObjectKeywords(_('Document'), document.number, document.keyword_list)
                        # Try to save the Document
                        self.save_document(document)
                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)
                        # See if the Document ID has been changed.  If it has, update the tree.
                        if document.id != originalName:
                            # Rename the Tree node
                            self.tree.SetItemText(sel, document.id)

                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # Begin constructing the message with the old and new names for the node
                                msg = " >|< %s >|< %s" % (originalName, document.id)
                                # Get the full Node Branch by climbing it to two levels above the root
                                while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                    # Update the selected node indicator
                                    sel = self.tree.GetItemParent(sel)
                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                    msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                # This must be untranslated to avoid problems in mixed-language environments.
                                # Prepend these on the Messsage
                                msg = "DocumentNode >|< Libraries" + msg
                                if DEBUG:
                                    print 'Message to send = "RN %s"' % msg
                                # Send the Rename Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            msg = 'Document %d' % document.number
                            if DEBUG:
                                print 'Message to send = "UKL %s"' % msg
                            if TransanaGlobal.chatWindow != None:
                                # Send the "Update Keyword List" message
                                TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Refresh the Keyword List, if it's a changed Keyword error
                    dlg.refresh_keywords()
                    # Highlight the first non-existent keyword in the Keywords control
                    dlg.highlight_bad_keyword()

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
            # Unlock the record regardless of what happens
            document.unlock_record()

    def save_document(self, document):
        """Save/Update the Document object."""
        if document != None:
            document.db_save()
       
    def add_episode(self, library_name):
        """User interface for adding a new episode."""
        # Load the Library which contains the Episode
        library = Library.Library(library_name)
        try:
            # Lock the Library, to prevent it from being deleted out from under the add.
            library.lock_record()
            libraryLocked = True
        # Handle the exception if the record is already locked by someone else
        except RecordLockedError, s:
            # If we can't get a lock on the Library, it's really not that big a deal.  We only try to get it
            # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
            # still be able to add Episodes even if someone else is editing the Library properties.
            libraryLocked = False
        # Create the Episode Properties Dialog Box to Add an Episode
        dlg = EpisodePropertiesForm.AddEpisodeDialog(self, -1, library)
        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Use "try", as exceptions could occur
            try:
                # Display the Episode Properties Dialog Box and get the data from the user
                episode = dlg.get_input()
                # If the user pressed OK ...
                if episode != None:
                    # Try to save the data from the form
                    self.save_episode(episode)
                    nodeData = (_('Libraries'), library.id, episode.id)
                    self.tree.add_Node('EpisodeNode', nodeData, episode.number, library.number)
                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        if DEBUG:
                            print 'Message to send = "AE %s >|< %s"' % (nodeData[-2], nodeData[-1])
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("AE %s >|< %s" % (nodeData[-2], nodeData[-1]))

                    # If we do all this, we don't need to continue any more.
                    contin = False
                    # Unlock the Library, if we locked it.
                    if libraryLocked:
                        library.unlock_record()
                    # Return the Episode Name so that the Transcript can be added to the proper Episode
                    return episode.id
                # If the user pressed Cancel ...
                else:
                    # ... then we don't need to continue any more.
                    contin = False
                    # Unlock the Library, if we locked it.
                    if libraryLocked:
                        library.unlock_record()
                    # If no Episode is created, indicate this so that no Transcript will be created.
                    return None
            # Handle "SaveError" exception
            except SaveError:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()

                # Refresh the Keyword List, if it's a changed Keyword error
                dlg.refresh_keywords()
                # Highlight the first non-existent keyword in the Keywords control
                dlg.highlight_bad_keyword()

            # Handle other exceptions
            except:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    
                # Display the Exception Message, allow "continue" flag to remain true
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
        
    def edit_episode(self, episode):
        """User interface for editing an episode."""
        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            episode.lock_record()
        # Handle the exception if the record is already locked by someone else
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Episode"), episode.id)
        # If the record is not locked, keep going.
        else:
            # Create the Episode Properties Dialog Box to edit the Episode Properties
            dlg = EpisodePropertiesForm.EditEpisodeDialog(self, -1, episode)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Episode Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # Check to see if there are keywords to be propagated
                        self.ControlObject.PropagateObjectKeywords(_('Episode'), episode.number, episode.keyword_list)
                        # Try to save the Episode
                        self.save_episode(episode)
                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)
                        # See if the Episode ID has been changed.  If it has, update the tree.
                        if episode.id != originalName:
                            # Rename the Tree node
                            self.tree.SetItemText(sel, episode.id)
                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # Begin constructing the message with the old and new names for the node
                                msg = " >|< %s >|< %s" % (originalName, episode.id)
                                # Get the full Node Branch by climbing it to two levels above the root
                                while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                    # Update the selected node indicator
                                    sel = self.tree.GetItemParent(sel)
                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                    msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                # This must be untranslated to avoid problems in mixed-language environments.
                                # Prepend these on the Messsage
                                msg = "EpisodeNode >|< Libraries" + msg
                                if DEBUG:
                                    print 'Message to send = "RN %s"' % msg
                                # Send the Rename Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            msg = 'Episode %d' % episode.number
                            if DEBUG:
                                print 'Message to send = "UKL %s"' % msg
                            if TransanaGlobal.chatWindow != None:
                                # Send the "Update Keyword List" message
                                TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Refresh the Keyword List, if it's a changed Keyword error
                    dlg.refresh_keywords()
                    # Highlight the first non-existent keyword in the Keywords control
                    dlg.highlight_bad_keyword()

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
            # Unlock the record regardless of what happens
            episode.unlock_record()

    def save_episode(self, episode):
        """Save/Update the Episode object."""
        # FIXME: Graceful exception handling
        if episode != None:
            episode.db_save()
       
    def add_transcript(self, library_name, episode_name):
        """User interface for adding a new transcript."""
        # Load the Episode Object that parents the Transcript
        episode = Episode.Episode(series=library_name, episode=episode_name)
        try:
            # Lock the Episode, to prevent it from being deleted out from under the add.
            episode.lock_record()
            episodeLocked = True
        # Handle the exception if the record is already locked by someone else
        except RecordLockedError, e:
            # If we can't get a lock on the Episode, it's really not that big a deal.  We only try to get it
            # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
            # still be able to add Transcripts even if someone else is editing the Episode properties.
            episodeLocked = False
        # Create the Transcript Properties Dialog Box to Add a Trancript
        dlg = TranscriptPropertiesForm.AddTranscriptDialog(self, -1, episode)
        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Use "try", as exceptions could occur
            try:
                # Display the Transcript Properties Dialog Box and get the data from the user
                transcript = dlg.get_input()
                # If the user pressed OK ...
                if transcript != None:
                    # Try to save the data from the form
                    self.save_transcript(transcript)
                    nodeData = (_('Libraries'), library_name, episode_name, transcript.id)
                    self.tree.add_Node('TranscriptNode', nodeData, transcript.number, episode.number)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        if DEBUG:
                            print 'Message to send = "AT %s >|< %s >|< %s"' % (nodeData[-3], nodeData[-2], nodeData[-1])
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("AT %s >|< %s >|< %s" % (nodeData[-3], nodeData[-2], nodeData[-1]))

                    # If we do all this, we don't need to continue any more.
                    contin = False
                    # Unlock the record
                    if episodeLocked:
                        episode.unlock_record()

                    # Note the original Tree Selection
                    sel = self.tree.GetSelections()[0]
                    # Sort the Transcript Node
                    self.tree.SortChildren(sel)

                    # return the Transcript ID so that it can be loaded
                    return transcript.id
                # If the user pressed Cancel ...
                else:
                    # ... then we don't need to continue any more.
                    contin = False
                    # Unlock the record
                    if episodeLocked:
                        episode.unlock_record()
                    # Returning None signals that the user cancelled, so no Transcript can be loaded.
                    return None
            # Handle "SaveError" exception
            except SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
            # Handle other exceptions
            except:
                # Display the Exception Message, allow "continue" flag to remain true
                prompt = "%s\n%s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % sys.exc_info()[:2])
                errordlg.ShowModal()
                errordlg.Destroy()

    def edit_transcript(self, transcript):
        """User interface for editing a transcript."""
        # use "try", as exceptions could occur
        try:
##            # If the user tries to edit the currently-loaded Transcript ...
##            if (transcript.number in self.ControlObject.TranscriptNum.keys()):
##                # Bring the Transcript Window to the front
##                self.ControlObject.BringTranscriptToFront()
##                # ... first clear all the windows ...
##                self.ControlObject.ClearAllWindows()
##                # ... then reload the transcript in case it got changed (saved) during the ClearAllWindows call.
##                transcript.db_load_by_num(transcript.number)

            # If the user tries to edit a currently-loaded Transcript, close it first.
            self.ControlObject.CloseOpenTranscriptWindowObject(Transcript.Transcript, transcript.number)
            # This could have let to a change in the transcript object!  We better re-load it just to be safe.
            transcript.db_load_by_num(transcript.number)
        
            # Try to get a Record Lock.
            transcript.lock_record()
            # If the record is not locked, keep going.
            # Create the Transcript Properties Dialog Box to edit the Transcript Properties
            dlg = TranscriptPropertiesForm.EditTranscriptDialog(self, -1, transcript)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Transcript Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        self.save_transcript(transcript)
                        # See if this affects the currently loaded
                        # Transcript, in which case we have to re-load the
                        # transcript if a new one was imported.
                        if transcript.number in self.ControlObject.TranscriptNum.keys():
                            self.ControlObject.LoadTranscript(transcript.series_id, transcript.episode_id, transcript.number)

                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)
                        # See if the Transcript ID has been changed.  If it has, update the tree.
                        if transcript.id != originalName:
                            # Rename the Tree node
                            self.tree.SetItemText(sel, transcript.id)
                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # Begin constructing the message with the old and new names for the node
                                msg = " >|< %s >|< %s" % (originalName, transcript.id)
                                # Get the full Node Branch by climbing it to two levels above the root
                                while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                    # Update the selected node indicator
                                    sel = self.tree.GetItemParent(sel)
                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                    msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                # This must be untranslated to avoid problems in mixed-language environments.
                                # Prepend these on the Messsage
                                msg = "TranscriptNode >|< Libraries" + msg
                                if DEBUG:
                                    print 'Message to send = "RN %s"' % msg
                                # Send the Rename Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()

                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    
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

                    import traceback
                    traceback.print_exc(file=sys.stdout)

            # Unlock the record regardless of what happens
            transcript.unlock_record()
        # Handle the exception if the record is locked
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Transcript"), transcript.id)

    def save_transcript(self, transcript):
        """Save/Update the Transcript object."""
        # FIXME: Graceful exception handling
        if transcript != None:
            transcript.db_save()

    def add_collection(self, ParentNum):
        """User interface for adding a new collection."""
        # If we're adding a nested Collection, we should lock the parent Collection
        if ParentNum > 0:
            # Get the Parent Collection
            collection = Collection.Collection(ParentNum)
            try:
                # Lock the Collectoin, to prevent it from being deleted out from under the add.
                collection.lock_record()
                collectionLocked = True
            # Handle the exception if the record is already locked by someone else
            except RecordLockedError, c:
                # If we can't get a lock on the Collection, it's really not that big a deal.  We only try to get it
                # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
                # still be able to add Nested Collections even if someone else is editing the Collection properties.
                collectionLocked = False
        else:
            collectionLocked = False
        # Create the Collection Properties Dialog Box to Add a Collection
        dlg = CollectionPropertiesForm.AddCollectionDialog(self, -1, ParentNum)
        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Use "try", as exceptions could occur
            try:
                # Display the Collection Properties Dialog Box and get the data from the user
                coll = dlg.get_input()
                # If the user pressed OK ...
                if coll != None:
                    # Try to save the data from the form
                    self.save_collection(coll)
                    nodeData = (_('Collections'),) + coll.GetNodeData()
                    # Add the new Collection to the data tree
                    self.tree.add_Node('CollectionNode', nodeData, coll.number, coll.parent)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "AC %s"
                        data = (nodeData[1],)
                        for nd in nodeData[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        if DEBUG:
                            print 'Message to send =', msg % data
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)
                    # Unlock the parent collection
                    if collectionLocked:
                        collection.unlock_record()
                    # If we do all this, we don't need to continue any more.
                    contin = False
                # If the user pressed Cancel ...
                else:
                    # Unlock the parent collection
                    if collectionLocked:
                        collection.unlock_record()
                    # ... then we don't need to continue any more.
                    contin = False

                # Note the original Tree Selection
                sel = self.tree.GetSelections()[0]
                # Sort the Collection Node
                self.tree.SortChildren(sel)

            # Handle "SaveError" exception
            except SaveError:
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
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
       
    def edit_collection(self, coll):
        """User interface for editing a collection."""
        # Assume failure unless proven otherwise
        success = False
        # Okay, when we're converting a Search Result to a Collection, we create a Collection named after the Search Results
        # and then use this routine.  That works well, except if there's a duplicate Collection ID in the Multi-user version,
        # the other client's collection will get renamed rather than the new collection being added.  Therefore, we need to
        # detect that this is what's going on and avoid that problem.  The following line detects that, as the Search Results
        # collection hasn't actually been saved yet.
        if coll.number == 0:
            fakeEdit = True
        else:
            fakeEdit = False
        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            coll.lock_record()
        # Handle the exception if the record is locked
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Collection"), coll.id)
        # If the record is not locked, keep going.
        else:
            # Create the Collection Properties Dialog Box to edit the Collection Properties
            dlg = CollectionPropertiesForm.EditCollectionDialog(self, -1, coll)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Collection Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # try to save the Collection
                        self.save_collection(coll)
                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)
                        # See if the Collection ID has been changed.  If it has, update the tree.
                        # That is, unless this is a Search Result conversion, as signalled by fakeEdit.
                        if (coll.id != originalName) and (not fakeEdit):
                            # Rename the Tree node
                            self.tree.SetItemText(sel, coll.id)
                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # Begin constructing the message with the old and new names for the node
                                msg = " >|< %s >|< %s" % (originalName, coll.id)
                                # Get the full Node Branch by climbing it to two levels above the root
                                while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                    # Update the selected node indicator
                                    sel = self.tree.GetItemParent(sel)
                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                    msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                # This must be untranslated to avoid problems in mixed-language environments.
                                # Prepend these on the Messsage
                                msg = "CollectionNode >|< Collections" + msg
                                if DEBUG:
                                    print 'Message to send = "RN %s"' % msg
                                # Send the Rename Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # If we get here, the save worked!
                        success = True
                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
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
            # Unlock the record regardless of what happens
            coll.unlock_record()
        # Return the "success" indicator to the calling routine
        return success

    def save_collection(self, collection):
        """Save/Update the Collection object."""
        # FIXME: Graceful exception handling
        if collection != None:
            collection.db_save()
 
    def add_quote(self, collection_id):
        """ User interface for adding a Quote to a Collection. """

        # Identify the selected Tree Node and its accompanying data
        sel = self.tree.GetSelections()[0]
        selData = self.tree.GetPyData(sel)

        # specify the data formats to accept.
        #   Our data could be a QuoteDragDropData object if the source is the Document (quote creation)
        dfQuote = wx.CustomDataFormat('QuoteDragDropData')

        # Specify the data object to accept data for these formats
        #   A QuoteDragDropData object will populate the cdoQuote object
        cdoQuote = wx.CustomDataObject(dfQuote)

        # Create a composite Data Object from the object types defined above
        cdo = wx.DataObjectComposite()
        cdo.Add(cdoQuote)
        # Open the Clipboard
        wx.TheClipboard.Open()
        # Try to get data from the Clipboard
        success = wx.TheClipboard.GetData(cdo)
        # If the data in the clipboard is in an appropriate format ...
        if success:
            # ... unPickle the data so it's in a usable form
            # Lets try to get a QuoteDragDropData object
            try:
                data2 = cPickle.loads(cdoQuote.GetData())
            except:
                # Windows doesn't fail this way, because success is never true above on Windows as it is on Mac
                data2 = None
                success = False
                # If this fails, that's okay
                pass

            # if we got a QuoteDragDataObject ...
            if isinstance(data2, DragAndDropObjects.QuoteDragDropData):
                # ... we should create a Quote!!!
                DragAndDropObjects.CreateQuote(data2, selData, self.tree, sel)
            else:
                # Mac fails this way
                success = False
            # Sort the Quotes, Clips, and Snapshots in the Collection
            self.tree.SortChildren(sel)

        # If there's not a Quote Creation object in the Clipboard, display an error message
        if not success:
            prompt = _("To add a quote, select the desired portion of a document and drag the selection onto a Collection.")
            dlg = Dialogs.InfoDialog(self, prompt)
            dlg.ShowModal()
            dlg.Destroy()
        # Close the Clipboard
        wx.TheClipboard.Close()

    def edit_quote(self, quote):
        """User interface for editing a quote."""

        # If the user tries to edit a currently-loaded Quote, close it first.
        self.ControlObject.CloseOpenTranscriptWindowObject(Quote.Quote, quote.number)
        # This could have let to a change in the Quote object!  We better re-load it just to be safe.
        quote.db_load_by_num(quote.number)
        
        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            quote.lock_record()
        # Handle the exception if the record is already locked by someone else
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Quote"), quote.id)
        # If the record is not locked, keep going.
        else:
            # For Quote Change Propagation, we actually need a copy of the Quote from before the user changes anything.
            # So let's make a copy of the Quote!
            originalQuote = quote.duplicate()
            # When you copy a quote, it has "0" for a Quote Number.  Let's preserve the original Quote Number.
            # (We just have to be careful NOT TO SAVE originalQuote!)
            originalQuote.number = quote.number
            # Create the Quote Properties Dialog Box to edit the Quote Properties
            dlg = QuotePropertiesForm.EditQuoteDialog(self, -1, quote)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Quote Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # Try to save the Quote
                        self.save_quote(quote)
                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)
                        # See if the Quote ID has been changed.  If it has, update the tree.
                        if quote.id != originalName:
                            # Rename the Tree node
                            self.tree.SetItemText(sel, quote.id)

                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # Begin constructing the message with the old and new names for the node
                                msg = " >|< %s >|< %s" % (originalName, quote.id)
                                # Get the full Node Branch by climbing it to two levels above the root
                                while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                    # Update the selected node indicator
                                    sel = self.tree.GetItemParent(sel)
                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                    msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                # This must be untranslated to avoid problems in mixed-language environments.
                                # Prepend these on the Messsage
                                msg = "QuoteNode >|< Collections" + msg
                                if DEBUG:
                                    print 'Message to send = "RN %s"' % msg
                                # Send the Rename Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # If the current main object is a Document and it's the Document that contains the
                        # edited Quote, we need to update the Keyword Visualization!
                        if (isinstance(self.ControlObject.currentObj, Document.Document)) and \
                           (quote.source_document_num == self.ControlObject.currentObj.number):
                            self.ControlObject.UpdateKeywordVisualization()
                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            msg = 'Quote %d' % quote.number
                            if DEBUG:
                                print 'Message to send = "UKL %s"' % msg
                            if TransanaGlobal.chatWindow != None:
                                # Send the "Update Keyword List" message
                                TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                            # If the Quote Properties Dialog was closed by pressing the Propagate Quote Changes button,
                            # we can detect that by looking at the propagatePressed property of the form.  Now that the
                            # Quote changes have been successfully saved, we need to deal with propagation if requested.
                            if dlg.propagatePressed:
                                # Start up the Propagate Changes tool, passing in the original Quote copy and the proper data
                                # from the edited Quote.
                                propagateDlg = PropagateChanges.PropagateClipChanges(self,
                                                                                     "Quote",
                                                                                     originalQuote,
                                                                                     -1,
                                                                                     quote.text,
                                                                                     quote.id,
                                                                                     quote.keyword_list)

                            # Even if this computer doesn't need to update the keyword visualization others, might need to.
                            if not TransanaConstants.singleUserVersion:
                                # We need to pass the type of the current object, the Quote's record number, and
                                # the Quote's Document number.
                                if DEBUG:
                                    print 'Message to send = "UKV %s %s %s"' % ('Quote', quote.number, quote.source_document_num)
                                    
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Quote', quote.number, quote.source_document_num))

                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Refresh the Keyword List, if it's a changed Keyword error
                    dlg.refresh_keywords()
                    # Highlight the first non-existent keyword in the Keywords control
                    dlg.highlight_bad_keyword()

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
            # Unlock the record regardless of what happens
            quote.unlock_record()

    def save_quote(self, quote):
        """Save/Update the Quote object."""
        if quote != None:
            quote.db_save()

    def add_clip(self, collection_id):
        """ User interface for adding a Clip to a Collection. """

        # Identify the selected Tree Node and its accompanying data
        sel = self.tree.GetSelections()[0]
        selData = self.tree.GetPyData(sel)

        # specify the data formats to accept.
        #   Our data could be a ClipDragDropData object if the source is the Transcript (clip creation)
        dfClip = wx.CustomDataFormat('ClipDragDropData')

        # Specify the data object to accept data for these formats
        #   A ClipDragDropData object will populate the cdoClip object
        cdoClip = wx.CustomDataObject(dfClip)

        # Create a composite Data Object from the object types defined above
        cdo = wx.DataObjectComposite()
        cdo.Add(cdoClip)
        # Open the Clipboard
        wx.TheClipboard.Open()
        # Try to get data from the Clipboard
        success = wx.TheClipboard.GetData(cdo)
        # If the data in the clipboard is in an appropriate format ...
        if success:
            # ... unPickle the data so it's in a usable form
            # Lets try to get a ClipDragDropData object
            try:
                data2 = cPickle.loads(cdoClip.GetData())
            except:
                # Windows doesn't fail this way, because success is never true above on Windows as it is on Mac
                data2 = None
                success = False
                # If this fails, that's okay
                pass

            # if we got a ClipDragDataObject ...
            if type(data2) == type(DragAndDropObjects.ClipDragDropData()):
                # ... we should create a Clip!!!
                DragAndDropObjects.CreateClip(data2, selData, self.tree, sel)
            else:
                # Mac fails this way
                success = False

        # Sort the Quotes, Clips, and Snapshots in the Collection
        self.tree.SortChildren(sel)

        # If there's not a Clip Creation object in the Clipboard, display an error message
        if not success:
            prompt = _("To add a Clip, select the desired portion of a transcript and drag the selection onto a Collection.")
            dlg = Dialogs.InfoDialog(self, prompt)
            dlg.ShowModal()
            dlg.Destroy()
        # Close the Clipboard
        wx.TheClipboard.Close()

    def edit_clip(self, clip):
        """User interface for editing a clip."""
        # If the user wants to edit the currently-loaded Clip ...
        if ((type(self.ControlObject.currentObj) == type(Clip.Clip())) and \
            (self.ControlObject.currentObj.number == clip.number)):
            # ... we should clear the clip from the interface before editing it.
            # Set the active transcript to signal the need to clear ALL windows ...
            self.ControlObject.activeTranscript = 0
            # ... and clear the Windows
            self.ControlObject.ClearAllWindows(clearAllPanes=True)
            # ... then reload the clip in case it got changed (saved) during the ClearAllWindows call.
            clip.db_load_by_num(clip.number)

        # Begin exception handling
        try:
            # If we have an existing clip ...
            if clip.number != 0:
                # Try to get a Record Lock
                clip.lock_record()
                # flag that it was an existing clip
                existingClip = True
            # If we don't have an existing clip ...
            else:
                # ... flag that the clips did not exist
                existingClip = False
                # If it has no Sort Order ...
                if clip.sort_order == 0:
                    # ... give it the highest Sort Order for the collection
                    clip.sort_order = DBInterface.getMaxSortOrder(clip.collection_num) + 1

        # Handle the exception if the record is locked
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Clip"), clip.id)
        # If the record is not locked, keep going.
        else:
            # Remember the clip's original name
            originalClipID = clip.id
            # For Clip Change Propagation, we actually need a copy of the clip from before the user changes anything.
            # So let's make a copy of the clip!
            originalClip = clip.duplicate()
            # When you copy a clip, it has "0" for a Clip Number.  Let's preserve the original Clip Number.
            # (We just have to be careful NOT TO SAVE originalClip!)
            originalClip.number = clip.number
            # Create the Clip Properties Dialog Box to edit the Clip Properties
            dlg = ClipPropertiesForm.EditClipDialog(self, -1, clip)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Clip Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # Create a Popup Dialog
                        tmpDlg = Dialogs.PopupDialog(None, _("Saving Clip"), _("Saving the Clip"))
                        # Try to save the Clip Data
                        self.save_clip(clip)
                        # If any keywords that served as Keyword Examples that got removed from the Clip,
                        # we need to remove them from the Database Tree.  
                        for (keywordGroup, keyword, clipNum) in dlg.keywordExamplesToDelete:
                            # Load the specified Clip record.  We can speed the load by not loading the Clip Transcript(s)
                            tempClip = Clip.Clip(clipNum, skipText=True)
                            # Prepare the Node List for removing the Keyword Example Node, using the clip's original name
                            nodeList = (_('Keywords'), keywordGroup, keyword, originalClipID)
                            # Call the DB Tree's delete_Node method.  Include the Clip Record Number so the correct Clip entry will be removed.
                            self.tree.delete_Node(nodeList, 'KeywordExampleNode', tempClip.number)
                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)
                        # if the clip previously existed ...
                        if existingClip:
                            # See if the Clip ID has been changed.  If it has, update the tree.
                            if clip.id != originalName:
                                for (kwg, kw, clipNumber, clipID) in DBInterface.list_all_keyword_examples_for_a_clip(clip.number):
                                    nodeList = (_('Keywords'), kwg, kw, self.tree.GetItemText(sel))
                                    exampleNode = self.tree.select_Node(nodeList, 'KeywordExampleNode')
                                    self.tree.SetItemText(exampleNode, clip.id)
                                    # If we're in the Multi-User mode, we need to send a message about the change
                                    if not TransanaConstants.singleUserVersion:
                                        # Begin constructing the message with the old and new names for the node
                                        msg = " >|< %s >|< %s" % (originalName, clip.id)
                                        # Get the full Node Branch by climbing it to two levels above the root
                                        while (self.tree.GetItemParent(self.tree.GetItemParent(exampleNode)) != self.tree.GetRootItem()):
                                            # Update the selected node indicator
                                            exampleNode = self.tree.GetItemParent(exampleNode)
                                            # Prepend the new Node's name on the Message with the appropriate seperator
                                            msg = ' >|< ' + self.tree.GetItemText(exampleNode) + msg
                                        # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                        # This must be untranslated to avoid problems in mixed-language environments.
                                        # Prepend these on the Messsage
                                        msg = "KeywordExampleNode >|< Keywords" + msg
                                        if DEBUG:
                                            print 'Message to send = "RN %s"' % msg
                                        # Send the Rename Node message
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)
                                            
                                # Rename the Tree node
                                self.tree.SetItemText(sel, clip.id)
                                # If we're in the Multi-User mode, we need to send a message about the change
                                if not TransanaConstants.singleUserVersion:
                                    # Begin constructing the message with the old and new names for the node
                                    msg = " >|< %s >|< %s" % (originalName, clip.id)
                                    # Get the full Node Branch by climbing it to two levels above the root
                                    while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                        # Update the selected node indicator
                                        sel = self.tree.GetItemParent(sel)
                                        # Prepend the new Node's name on the Message with the appropriate seperator
                                        msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                    # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                    # This must be untranslated to avoid problems in mixed-language environments.
                                    # Prepend these on the Messsage
                                    msg = "ClipNode >|< Collections" + msg
                                    if DEBUG:
                                        print 'Message to send = "RN %s"' % msg
                                    # Send the Rename Node message
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                msg = 'Clip %d' % clip.number
                                if DEBUG:
                                    print 'Message to send = "UKL %s"' % msg
                                if TransanaGlobal.chatWindow != None:
                                    # Send the "Update Keyword List" message
                                    TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
                            # If the Clip Properties Dialog was closed by pressing the Propagate Clip Changes button,
                            # we can detect that by looking at the propagatePressed property of the form.  Now that the
                            # clip changes have been successfully saved, we need to deal with propagation if requested.
                            if dlg.propagatePressed:
                                # Close popup dialog prior to showing the Propagate Clip Changes form
                                tmpDlg.Destroy()
                                # Signal that the tmpDlg has been closed.
                                tmpDlg = None
                                # Start up the Propagate Changes tool, passing in the original clip copy and the proper data
                                # from the edited clip.
                                propagateDlg = PropagateChanges.PropagateClipChanges(self,
                                                                                     "Clip",
                                                                                     originalClip,
                                                                                     -1,
                                                                                     clip.transcripts,
                                                                                     clip.id,
                                                                                     clip.keyword_list)

                            # If we do all this, we don't need to continue any more.
                            contin = False
                        # If the clip did not previously exist ...
                        else:
                            # Get the node data
                            nodeData = (_('Collections'),) + clip.GetNodeData()
                            # Add the new Collection to the data tree
                            self.tree.add_Node('ClipNode', nodeData, clip.number, clip.collection_num, sortOrder=clip.sort_order)

                            try:
                                wx.Yield()
                            except:
                                pass

                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                msg = "ACl %s"
                                data = (nodeData[1],)
                                for nd in nodeData[2:]:
                                    msg += " >|< %s"
                                    data += (nd, )
                                if DEBUG:
                                    print 'Message to send =', msg % data
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage(msg % data)


                            # If we do all this, we don't need to continue any more.
                            contin = False

                        # If the current main object is an Episode and it's the episode that contains the
                        # edited Clip, we need to update the Keyword Visualization!
                        if (isinstance(self.ControlObject.currentObj, Episode.Episode)) and \
                           (clip.episode_num == self.ControlObject.currentObj.number):
                            self.ControlObject.UpdateKeywordVisualization()
                        # Even if this computer doesn't need to update the keyword visualization others, might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to pass the type of the current object, the Clip's record number, and
                            # the Clip's Episode number.
                            if DEBUG:
                                print 'Message to send = "UKV %s %s %s"' % ('Clip', clip.number, clip.episode_num)
                                
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', clip.number, clip.episode_num))

                        # If the tmpDlg is still open ...
                        if tmpDlg is not None:
                            # Close popup dialog
                            tmpDlg.Destroy()
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
                    # If the tmpDlg is still open ...
                    if tmpDlg is not None:
                        # Close popup dialog
                        tmpDlg.Destroy()
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()

                    # Refresh the Keyword List, if it's a changed Keyword error
                    dlg.refresh_keywords()
                    # Highlight the first non-existent keyword in the Keywords control
                    dlg.highlight_bad_keyword()

                # Handle other exceptions
                except:
                    # If the tmpDlg is still open ...
                    if tmpDlg is not None:
                        # Close popup dialog
                        tmpDlg.Destroy()
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
            # If we had an existing clip ...
            if existingClip:
                # Unlock the record regardless of what happens
                clip.unlock_record()

    def save_clip(self, clip):
        """Save/Update the Clip object."""
        # FIXME: Graceful exception handling
        if clip != None:
            clip.db_save()

    def add_snapshot(self, parentNum):
        """ User Interface for adding a Snapshot to a Collection """
        # Identify the selected Tree Node and its accompanying data
        sel = self.tree.GetSelections()[0]
        selData = self.tree.GetPyData(sel)

        tempCollection = Collection.Collection(parentNum)
        try:
            # Lock the parent Collection, to prevent it from being deleted out from under the add.
            tempCollection.lock_record()
            collectionLocked = True
        # Handle the exception if the record is already locked by someone else
        except TransanaExceptions.RecordLockedError, c:
            # If we can't get a lock on the Collection, it's really not that big a deal.  We only try to get it
            # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
            # still be able to add Snapshots even if someone else is editing the Collection properties.
            collectionLocked = False

        # Create a Snapshot Object
        tmpSnapshot = Snapshot.Snapshot()
        # Specify the Parent Collection
        tmpSnapshot.collection_num = parentNum
        # Create the Snapshot Properties Dialog Box to Add a Snapshot
        dlg = SnapshotPropertiesForm.AddSnapshotDialog(self, -1, tmpSnapshot)

        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Use "try", as exceptions could occur
            try:
                # Display the Snapshot Properties Dialog Box and get the data from the user
                tmpSnapshot = dlg.get_input()
                # If the user pressed OK ...
                if tmpSnapshot != None:
                    # If we're on a Collection Node and the Sort Order is the default (0) ...
                    if (selData.nodetype == 'CollectionNode') and (tmpSnapshot.sort_order == 0):
                        # ... then put this Snapshot at the end of the list
                        tmpSnapshot.sort_order = DBInterface.getMaxSortOrder(tmpSnapshot.collection_num) + 1
                        # ... and we don't need an InsertPos
                        insertPos = None
                    # If we're on a Quote, Clip or Snapshot ...
                    elif selData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
                        # ... then we pass the node as the InsertPos
                        insertPos = sel
                    # Try to save the data from the form
                    self.save_snapshot(tmpSnapshot)
                    # Set up the Nodes for the new Shapshot
                    nodeData = (_('Collections'),) +  tmpSnapshot.GetNodeData(True)
                    # Add the Snapshot to the Database Tree
                    self.tree.add_Node('SnapshotNode', nodeData, tmpSnapshot.number, tmpSnapshot.collection_num, sortOrder=tmpSnapshot.sort_order, insertPos=insertPos)

                    # If the Snapshot is connect to an Episode ...
                    if tmpSnapshot.episode_num != 0:
                        # See if the Keyword visualization needs to be updated.
                        self.ControlObject.UpdateKeywordVisualization()

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Create an "Add Snapshot" message
                        msg = "ASnap %s"
                        # Build the message details
                        data = (nodeData[1],)
                        for nd in nodeData[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        if DEBUG:
                            print 'Message to send =', msg % data
                        # If there's a Chat window ...
                        if TransanaGlobal.chatWindow != None:
                            # ... send the message
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                        # If the Snapshot is connect to an Episode ...
                        if tmpSnapshot.episode_num != 0:
                            # ... and there is a Chat Window ...
                            if TransanaGlobal.chatWindow != None:
                                # ... send a message to update the Keyword Visualization for the correct Episode
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', tmpSnapshot.episode_num, 0))

                    # If we've dropped on a Quote, Clip or Snapshot and the new Snapshot doesn't have a SortOrder ...
                    if (selData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']) and \
                       (tmpSnapshot.sort_order == 0):
                        # ... get the Collection where the Snapshot was created
                        tempCollection = Collection.Collection(tmpSnapshot.collection_num)
                        # Now change the Sort Order, and if it succeeds ...
                        if DragAndDropObjects.ChangeClipOrder(self.tree, sel, tmpSnapshot, tempCollection):
                            # ... let's send a message telling others they need to re-order this collection!
                            if not TransanaConstants.singleUserVersion:
                                # Start by getting the Collection's Node Data
                                nodeData = tempCollection.GetNodeData()
                                # Indicate that we're sending a Collection Node
                                msg = "CollectionNode"
                                # Iterate through the nodes
                                for node in nodeData:
                                    # ... add the appropriate seperator and then the node name
                                    msg += ' >|< %s' % node

                                if DEBUG:
                                    print 'Message to send = "OC %s"' % msg

                                # Send the Order Collection Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("OC %s" % msg)

                        # We're getting a "Last Save Time" error in the SnapshotWindow when entering Edit mode.
                        # Locking and unlocking here will prevent that by updating the object's LastSaveTime.
                        tmpSnapshot.lock_record()
                        tmpSnapshot.unlock_record()

                    # Load the Snapshot Interface
                    self.ControlObject.LoadSnapshot(tmpSnapshot)    # Load everything via the ControlObject

                # If we get this far, we're DONE.  Signal that we can exit the loop!
                contin = False

                # Sort the Quotes, Clips, and Snapshots in the Collection
                self.tree.SortChildren(sel)

            # Handle "SaveError" exception
            except SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()

                # Refresh the Keyword List, if it's a changed Keyword error
                dlg.refresh_keywords()
                # Highlight the first non-existent keyword in the Keywords control
                dlg.highlight_bad_keyword()

            # Handle other exceptions
            except:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)
                    
                # Display the Exception Message, allow "continue" flag to remain true
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
            
        # Unlock the parent collection
        if collectionLocked:
            tempCollection.unlock_record()
        # Destroy the dialog
        dlg.Destroy()

    def edit_snapshot(self, snapshot):
        """ User Interface for Editing a Snapshot's Properties """
        # If the user wants to edit properties for a Snapshot Window that is currently open, we should close the Snapshot!
        # Iterate through all open Snapshot Windows
        for snapshotWindow in self.ControlObject.SnapshotWindows:
            # If we find the desired Snapshot window already open ...
            if snapshotWindow.obj.number == snapshot.number:
                # ... close the Snapshot Window
                snapshotWindow.Close()
            # ... then reload the snapshot in case it got changed (saved) during the Close() call.
            snapshot.db_load(snapshot.number)

        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            snapshot.lock_record()
        # Handle the exception if the record is locked
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Snapshot"), snapshot.id)
        # If the record is not locked, keep going.
        else:
            # Create the Snapshot Properties Dialog Box to edit the Snapshot Properties
            dlg = SnapshotPropertiesForm.EditSnapshotDialog(self, -1, snapshot)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Snapshot Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # try to save the Collection
                        self.save_snapshot(snapshot)
                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)
                        # See if the Snapshot ID has been changed.  If it has, update the tree.
                        if (snapshot.id != originalName):
                            # Rename the Tree node
                            self.tree.SetItemText(sel, snapshot.id)
                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # Begin constructing the message with the old and new names for the node
                                msg = " >|< %s >|< %s" % (originalName, snapshot.id)
                                # Get the full Node Branch by climbing it to two levels above the root
                                while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                    # Update the selected node indicator
                                    sel = self.tree.GetItemParent(sel)
                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                    msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                # This must be untranslated to avoid problems in mixed-language environments.
                                # Prepend these on the Messsage
                                msg = "SnapshotNode >|< Collections" + msg
                                if DEBUG:
                                    print 'Message to send = "RN %s"' % msg
                                # Send the Rename Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # See if the Keyword visualization needs to be updated.
                        self.ControlObject.UpdateKeywordVisualization()
                        # Even if this computer doesn't need to update the keyword visualization others, might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to update the Episode Keyword Visualization
                            if DEBUG:
                                print 'Message to send = "UKV %s %s %s"' % ('Episode', snapshot.episode_num, 0)
                                
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', snapshot.episode_num, 0))

                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
                    # Display the Error Message, allow "continue" flag to remain true
                    errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                    errordlg.ShowModal()
                    errordlg.Destroy()

                    # Refresh the Keyword List, if it's a changed Keyword error
                    dlg.refresh_keywords()
                    # Highlight the first non-existent keyword in the Keywords control
                    dlg.highlight_bad_keyword()

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
            # Unlock the record regardless of what happens
            snapshot.unlock_record()


    def save_snapshot(self, snapshot):
        """ Save / Update the Snapshot object """
        # If we have a defined Snapshot object ...
        if snapshot != None:
            # ... save it!
            snapshot.db_save()

    def add_note(self, libraryNum=0, episodeNum=0, transcriptNum=0, collectionNum=0, clipNum=0, snapshotNum=0, documentNum=0, quoteNum=0):
        """User interface method for adding a Note to an object."""
        # We should lock the parent object when adding a note.
        # First, get the appropriate parent object
        if libraryNum > 0:
            # Load the Library Object
            parentObj = Library.Library(libraryNum)
            objType = _('Libraries')
        elif episodeNum > 0:
            # Load the Episode Object
            parentObj = Episode.Episode(episodeNum)
            objType = _("Episode")
        elif transcriptNum > 0:
            # Load the Transcript Object
            # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
            parentObj = Transcript.Transcript(transcriptNum, skipText=True)
            objType = _("Transcript")
        elif collectionNum > 0:
            # Load the Collection Object
            parentObj = Collection.Collection(collectionNum)
            objType = _("Collection")
        elif clipNum > 0:
            # Load the Clip Object.  We can speed the load by not loading the Clip Transcript(s)
            parentObj = Clip.Clip(clipNum)
            objType = _("Clip")
        elif snapshotNum > 0:
            # Load the Snapshot Object.
            parentObj = Snapshot.Snapshot(snapshotNum)
            objType = _("Snapshot")
        elif documentNum > 0:
            # Load the Document Object
            parentObj = Document.Document(documentNum)
            objType = _("Document")
        elif quoteNum > 0:
            # Load the Quote Object
            parentObj = Quote.Quote(quoteNum)
            objType = _("Quote")
        try:
            # Lock the parent object, to prevent it from being deleted out from under the add.
            parentObj.lock_record()

            # Create the Note Properties Dialog Box to Add a Note
            dlg = NotePropertiesForm.AddNoteDialog(self, -1, libraryNum, episodeNum, transcriptNum, collectionNum, clipNum, snapshotNum, documentNum, quoteNum)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # While the "continue" flag is True ...
            while contin:
                # Use "try", as exceptions could occur
                try:
                    # Display the Note Properties Dialog Box and get the data from the user
                    note = dlg.get_input()
                    # If the user pressed OK ...
                    if note != None:
                        # Create the Note Editing Form
                        noteedit = NoteEditor.NoteEditor(self, note.text)
                        # Display the Node Editing Form and get the user's input
                        note.text = noteedit.get_text()
                        # Try to save the data from the forms
                        self.save_note(note)

                        # Get the note's correct data
                        if libraryNum != 0:
                            nodeType = 'LibraryNoteNode'
                            library = parentObj
                            nodeData = ('Libraries', library.id, note.id)
                            parentNum = library.number
                            msgType = 'ASN'
                            nodeDataRoot = 'Libraries'
                        elif episodeNum != 0:
                            nodeType = 'EpisodeNoteNode'
                            episode = parentObj
                            nodeData = ('Libraries', episode.series_id, episode.id, note.id)
                            parentNum = episode.number
                            msgType = 'AEN'
                            nodeDataRoot = 'Libraries'
                        elif transcriptNum != 0:
                            nodeType = 'TranscriptNoteNode'
                            transcript = parentObj
                            episode = Episode.Episode(num = transcript.episode_num)
                            nodeData = ('Libraries', episode.series_id, episode.id, transcript.id, note.id)
                            parentNum = transcript.number
                            msgType = 'ATN'
                            nodeDataRoot = 'Libraries'
                        elif collectionNum != 0:
                            nodeType = 'CollectionNoteNode'
                            collection = parentObj
                            nodeData = ('Collections',) + collection.GetNodeData() + (note.id,)
                            parentNum = collection.number  # This is the NOTE's parent, not the Collection's parent!!
                            msgType = 'ACN'
                            nodeDataRoot = 'Collections'
                        elif clipNum != 0:
                            nodeType = 'ClipNoteNode'
                            clip = parentObj
                            collection = Collection.Collection(clip.collection_num)
                            nodeData = ('Collections',) + collection.GetNodeData() + (clip.id, note.id)
                            parentNum = clip.number
                            msgType = 'AClN'
                            nodeDataRoot = 'Collections'
                        elif snapshotNum != 0:
                            nodeType = 'SnapshotNoteNode'
                            snapshot = parentObj
                            collection = Collection.Collection(snapshot.collection_num)
                            nodeData = ('Collections',) + collection.GetNodeData() + (snapshot.id, note.id)
                            parentNum = snapshot.number
                            msgType = 'ASnN'
                            nodeDataRoot = 'Collections'
                        elif documentNum != 0:
                            nodeType = 'DocumentNoteNode'
                            document = parentObj
                            library = Library.Library(document.library_num)
                            nodeData = ('Libraries', library.id, document.id, note.id)
                            parentNum = document.number
                            msgType = 'ADN'
                            nodeDataRoot = 'Libraries'
                        elif quoteNum != 0:
                            nodeType = 'QuoteNoteNode'
                            quote = parentObj
                            collection = Collection.Collection(quote.collection_num)
                            nodeData = ('Collections',) + collection.GetNodeData() + (quote.id, note.id)
                            parentNum = quote.number
                            msgType = 'AQN'
                            nodeDataRoot = 'Collections'
                        else:
                            errordlg = Dialogs.ErrorDialog(None, 'Not Yet Implemented in DatabaseTreeTab.add_note()')
                            errordlg.ShowModal()
                            errordlg.Destroy()

                        # Add the Note to the Database Tree
                        self.tree.add_Node(nodeType, nodeData, note.number, parentNum)
                        # If the Notes Browser is open ...
                        if self.ControlObject.NotesBrowserWindow != None:
                            # ... add the new note to the Notes Browser
                            self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('A', note)
                        
                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            # Construct the message and data to be passed
                            msg = msgType + " %s"
                            # To avoid problems in mixed-language environments, we need the UNTRANSLATED string here!
                            data = (nodeDataRoot,)
                            for nd in nodeData[1:]:
                                msg += " >|< %s"
                                data += (nd, )
                                
                            if DEBUG:
                                print 'Message to send =', msg % data
                                
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage(msg % data)

                    # Unlock the parent Object
                    parentObj.unlock_record()

                    # Note the original Tree Selection
                    sel = self.tree.GetSelections()[0]
                    # Sort the Transcript Node
                    self.tree.SortChildren(sel)

                    # If we do all this, we don't need to continue any more.
                    contin = False
                # Handle "SaveError" exception
                except SaveError:
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
                    prompt = "%s : %s"
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(prompt, 'utf8')
                    errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                    errordlg.ShowModal()
                    errordlg.Destroy()
        # Handle the exception if the record is already locked by someone else
        except RecordLockedError, o:
            self.handle_locked_record(o, objType, parentObj.id)

    def edit_note(self, note):
        """User interface for editing a note."""
        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            note.lock_record()
        # Handle the exception if the record is locked
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Note"), note.id)
        # If the record is not locked, keep going.
        else:
            # Create the Note Properties Dialog Box to edit the Note Properties
            dlg = NotePropertiesForm.EditNoteDialog(self, -1, note)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # Note the original Tree Selection
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Note Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # Try to save the Note Data
                        self.save_note(note)
                        # Note the original name of the tree node
                        originalName = self.tree.GetItemText(sel)                        
                        # See if the Note ID has been changed.  If so ...
                        if note.id != originalName:
                            # Update the Tree node
                            self.tree.SetItemText(sel, note.id)
                            # If the Notes Browser is open ...
                            if self.ControlObject.NotesBrowserWindow != None:
                                # ... update the note in the Notes Browser
                                self.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('R', note, oldName=originalName)
                            # If we're in the Multi-User mode, we need to send a message about the change
                            if not TransanaConstants.singleUserVersion:
                                # Begin constructing the message with the old and new names for the node
                                msg = " >|< %s >|< %s" % (originalName, note.id)
                                # Get the full Node Branch by climbing it to two levels above the root
                                while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                    # Update the selected node indicator
                                    sel = self.tree.GetItemParent(sel)
                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                    msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                # For Notes, we need to know which root to climb, but we need the UNTRANSLATED root.
                                if 'unicode' in wx.PlatformInfo:
                                    libraryPrompt = unicode(_('Libraries'), 'utf8')
                                    collectionsPrompt = unicode(_('Collections'), 'utf8')
                                else:
                                    # We need the nodeType as the first element.  Then, 
                                    # we need the UNTRANSLATED label for the root node to avoid problems in mixed-language environments.
                                    libraryPrompt = _('Libraries')
                                    collectionsPrompt = _('Collections')
                                if self.tree.GetItemText(self.tree.GetItemParent(sel)) == libraryPrompt:
                                    rootNodeType = 'Libraries'
                                elif self.tree.GetItemText(self.tree.GetItemParent(sel)) == collectionsPrompt:
                                    rootNodeType = 'Collections'
                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                # This must be untranslated to avoid problems in mixed-language environments.
                                # Prepend these on the Messsage
                                msg = note.notetype + "NoteNode >|< " + rootNodeType + msg
                                if DEBUG:
                                    print 'Message to send = "RN %s"' % msg
                                # Send the Rename Node message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
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
            # Unlock the record regardless of what happens
            note.unlock_record()

    def save_note(self, note):
        """Save/Update the Note object."""
        # FIXME: Graceful exception handling
        if note != None:
            note.db_save()

    def add_keyword(self, keywordGroup):
        """User interface for adding a new keyword."""
        # Create the Keyword Properties Dialog Box to Add a Keyword
        dlg = KeywordPropertiesForm.AddKeywordDialog(self, -1, keywordGroup)
        # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
        contin = True
        # While the "continue" flag is True ...
        while contin:
            # Use "try", as exceptions could occur
            try:
                # Display the Keyword Properties Dialog Box and get the data from the user
                kw = dlg.get_input()
                # If the user pressed OK ...
                if kw != None:
                    # Try to save the data from the form
                    self.save_keyword(kw)
                    # Add the new Keyword to the tree
                    self.tree.add_Node('KeywordNode', (_('Keywords'), kw.keywordGroup, kw.keyword), 0, kw.keywordGroup)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        if DEBUG:
                            print 'Message to send = "AK %s >|< %s"' % (kw.keywordGroup, kw.keyword)
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (kw.keywordGroup, kw.keyword))

                    # If we do all this, we don't need to continue any more.
                    contin = False
                # If the user pressed Cancel ...
                else:
                    # ... then we don't need to continue any more.
                    contin = False

                # Note the original Tree Selection
                sel = self.tree.GetSelections()[0]
                # Sort the Transcript Node
                self.tree.SortChildren(sel)

            # Handle "SaveError" exception
            except SaveError:
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
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
        
    def edit_keyword(self, kw):
        """User interface for editing a keyword."""
        # use "try", as exceptions could occur
        try:
            # Try to get a Record Lock
            kw.lock_record()
        # Handle the exception if the record is locked
        except RecordLockedError, e:
            self.handle_locked_record(e, _("Keyword"), kw.keywordGroup + ':' + kw.keyword)
        # If the record is not locked, keep going.
        else:
            # Create the Keyword Properties Dialog Box to edit the Keyword Properties
            dlg = KeywordPropertiesForm.EditKeywordDialog(self, -1, kw)
            # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
            contin = True
            # See if the Keyword Group or Keyword has been changed.  If it has, update the tree.
            sel = self.tree.GetSelections()[0]
            # While the "continue" flag is True ...
            while contin:
                # if the user pressed "OK" ...
                try:
                    # Display the Keyword Properties Dialog Box and get the data from the user
                    if dlg.get_input() != None:
                        # Try to save the Keyword Data.  If this method returns True, we need to update the Database Tree to the new value.
                        if self.save_keyword(kw):
                            if kw.keywordGroup != self.tree.GetItemText(self.tree.GetItemParent(sel)):
                                # If the Keyword Group has changed, delete the Keyword Node, insert the new
                                # keyword group node if necessary, and insert the keyword in the right keyword
                                # group node
                                # Remove the old Keyword from the Tree
                                self.tree.delete_Node((_('Keywords'), self.tree.GetItemText(self.tree.GetItemParent(sel)), self.tree.GetItemText(sel)), 'KeywordNode')
                                # Add the new Keyword to the tree
                                self.tree.add_Node('KeywordNode', (_('Keywords'), kw.keywordGroup, kw.keyword), 0, kw.keywordGroup)

                                # Now let's communicate with other Transana instances if we're in Multi-user mode
                                if not TransanaConstants.singleUserVersion:
                                    if DEBUG:
                                        print 'Message to send = "AK %s >|< %s"' % (kw.keywordGroup, kw.keyword)
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (kw.keywordGroup, kw.keyword))

                            elif kw.keyword != self.tree.GetItemText(sel):
                                originalName = self.tree.GetItemText(sel)
                                # If only the Keyword has changed, simply rename the node
                                self.tree.SetItemText(sel, kw.keyword)
                                # If we're in the Multi-User mode, we need to send a message about the change
                                if not TransanaConstants.singleUserVersion:
                                    # Begin constructing the message with the old and new names for the node
                                    msg = " >|< %s >|< %s" % (originalName, kw.keyword)
                                    # Get the full Node Branch by climbing it to two levels above the root
                                    while (self.tree.GetItemParent(self.tree.GetItemParent(sel)) != self.tree.GetRootItem()):
                                        # Update the selected node indicator
                                        sel = self.tree.GetItemParent(sel)
                                        # Prepend the new Node's name on the Message with the appropriate seperator
                                        msg = ' >|< ' + self.tree.GetItemText(sel) + msg
                                    # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                    # This must be untranslated to avoid problems in mixed-language environments.
                                    # Prepend these on the Messsage
                                    msg = "KeywordNode >|< Keywords" + msg
                                    if DEBUG:
                                        print 'Message to send = "RN %s"' % msg
                                    # Send the Rename Node message
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)
                        # If the keyword save returns False, that signals a keyword merge, so we need to delete the tree node!
                        else:
                            keywordGroup = self.tree.GetItemText(self.tree.GetItemParent(sel))
                            keyword = self.tree.GetItemText(sel)
                            self.tree.delete_Node((_('Keywords'), keywordGroup, keyword), 'KeywordNode')
                        # If we do all this, we don't need to continue any more.
                        contin = False
                    # If the user pressed Cancel ...
                    else:
                        # ... then we don't need to continue any more.
                        contin = False
                # Handle "SaveError" exception
                except SaveError:
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
            # Unlock the record regardless of what happens
            kw.unlock_record()


    def save_keyword(self, kw):
        """Save/Update the Keyword object."""
        # FIXME: Graceful exception handling
        if kw != None:
            # We need to pass the keyword's DB_SAVE function result back to the calling routine.
            return kw.db_save()

    def GetSelectedNodeInfo(self):
        """ Get the Name, Record Number, Parent Number, and Type of the currently-selected Node """
        # Query the tree for the requested data and return it
        return self.tree.GetSelectedNodeInfo()

    def handle_locked_record(self, e, rtype, id):
        """Handle the RecordLockedError exception."""
        ReportRecordLockedException(rtype, id, e)


class MenuIDError(exceptions.Exception):
    """ Exception class for handling Menu ID Errors """
    def __init__(self, id=-1, menu=""):
        """ Initialize the MenuIDError Exception """
        # Create the appropriate error message
        if id > -1:
            msg = "Unable to handle selection menu ID %d for '%s' menu" % (id, menu)
        else:
            msg = "Unable to handle menu ID selection"
        # Assign the error message to the exception's arguments.
        self.explanation = msg


class _NodeData:
    """ This class defines the information that is known about each node in the Database Tree. """

    # NOTE:  _NodeType and DataTreeDragDropData have very similar structures so that they can be
    #        used interchangably.  If you alter one, please also alter the other.
   
    def __init__(self, nodetype='Unknown', recNum=0, parent=0, sortOrder=None, sourceObj=0):
        """ Initialize the NodeData Object """
        self.nodetype = nodetype    # nodetype indicates what sort of node we have.  Options include:
                                    # Root, LibraryRootNode, LibraryNode, DocumentNode, EpisodeNode, TranscriptNode,
                                    # CollectionsRootNode, CollectionNode, QuoteNode, ClipNode, SnapshotNode
                                    # KeywordsRootNode, KeywordGroupNode, KeywordNode, KeywordExampleNode
                                    # NotesGroupNode, LibraryNoteNode, EpisodeNoteNode, TranscriptNoteNode,
                                    # CollectionNoteNode, ClipNoteNode, SnapshotNoteNode, DocumentNoteNode, QuoteNodeNode, 
                                    # SearchRootNode, SearchResultsNode, SearchLibraryNode, SearchDocumentNode, SearchEpisodeNode,
                                    # SearchTranscriptNode, SearchCollectionNode, SearchQuoteNode, SearchClipNode, SearchSnapshotNode
        self.recNum = recNum        # recNum indicates the Database Record Number of the node
        self.parent = parent        # parent indicates the parent Record Number for nested Collections
        self.sortOrder = sortOrder  # sortOrder indicates order of Clips and Snapshots in a Collection
        self.sourceObj = sourceObj

    def __repr__(self):
        """ Provides a string representation of the data in the _NodeData object """
        str = 'nodetype = %s, recNum = %s, parent = %s, sourceObj = %s, sortOrder = %s' % (self.nodetype, self.recNum, self.parent, self.sourceObj, self.sortOrder)
        return str


class _DBTreeCtrl(wx.TreeCtrl):
    """Private class that implements the details of the tree widget."""
    def __init__(self, parent, id, pos, size, style):
        """ Initialize the Database Tree """
        wx.TreeCtrl.__init__(self, parent, id, pos, size, style)

        self.cmd_id = TransanaConstants.DATA_MENU_CMD_OFSET
        self.parent = parent
        
        # Track the number of Searches that have been requested
        self.searchCount = 1

        # Create image list of 16x16 object icons
        self.icon_list = ["Clip16", "Collection16", "Document16", "Episode16", "Keyword16",
                        "KeywordGroup16", "KeywordRoot16", 'Library16', 'LibraryRoot16', "Note16",
                        "NoteNode16", "Quote16", "SearchRoot16", "Search16",
                        "Snapshot16", "Transcript16", "db", "db_locked", "db_unlocked"]
        self.image_list = wx.ImageList(16, 16, 0, len(self.icon_list))
        self.image_list.Add(TransanaImages.Clip16.GetBitmap())
        self.image_list.Add(TransanaImages.Collection16.GetBitmap())
        self.image_list.Add(TransanaImages.Document16.GetBitmap())
        self.image_list.Add(TransanaImages.Episode16.GetBitmap())
        self.image_list.Add(TransanaImages.Keyword16.GetBitmap())
        self.image_list.Add(TransanaImages.KeywordGroup16.GetBitmap())
        self.image_list.Add(TransanaImages.KeywordRoot16.GetBitmap())
        self.image_list.Add(TransanaImages.Library16.GetBitmap())
        self.image_list.Add(TransanaImages.LibraryRoot16.GetBitmap())
        self.image_list.Add(TransanaImages.Note16.GetBitmap())
        self.image_list.Add(TransanaImages.NoteNode16.GetBitmap())
        self.image_list.Add(TransanaImages.Quote16.GetBitmap())
        self.image_list.Add(TransanaImages.SearchRoot16.GetBitmap())
        self.image_list.Add(TransanaImages.Search16.GetBitmap())
        self.image_list.Add(TransanaImages.Snapshot16.GetBitmap())
        self.image_list.Add(TransanaImages.Transcript16.GetBitmap())
        self.image_list.Add(TransanaImages.db.GetBitmap())
        self.image_list.Add(TransanaImages.db_locked.GetBitmap())
        self.image_list.Add(TransanaImages.db_unlocked.GetBitmap())
        self.SetImageList(self.image_list)

        # Define the Drop Target for the tree.  The custom drop target object
        # accepts the tree as a parameter so it can query it about where the
        # drop is supposed to be occurring to see if it will allow the drop.
        dt = DragAndDropObjects.DataTreeDropTarget(self)
        self.SetDropTarget(dt)
        
        self.refresh_tree()

        self.create_menus()

        # Initialize the cutCopyInfo Dictionary to empty values.
        # This information is used to facilitate Cut, Copy, and Paste functionality.
        self.cutCopyInfo = {'action': 'None', 'sourceItem': None, 'destItem': None}

        # Define the Begin Drag Event for the Database Tree Tab.  This enables Drag and Drop.
        wx.EVT_TREE_BEGIN_DRAG(self, id, self.OnCutCopyBeginDrag)

        # wx.EVT_MOTION(self, self.OnMotion)

        # Right Down initiates popup menu
        wx.EVT_RIGHT_DOWN(self, self.OnRightDown)

        # This processes double-clicks in the Tree Control
        wx.EVT_TREE_ITEM_ACTIVATED(self, id, self.OnItemActivated)

        # Prevent the ability to Edit Node Labels unless it is in a Search Result
        # or is explicitly processed in OnEndLabelEdit()
        wx.EVT_TREE_BEGIN_LABEL_EDIT(self, id, self.OnBeginLabelEdit)

        # Process Database Tree Label Edits
        wx.EVT_TREE_END_LABEL_EDIT(self, id, self.OnEndLabelEdit)

        # Process Key Presses
        self.Bind(wx.EVT_KEY_DOWN, self.OnKeyDown)

    def OnCompareItems(self, item1, item2):
        """ This method over-rides the wxTreeCtrl method, and implements the sort order for the TreeCtrl's SortChildren() method """
        # Get the PyData for the two nodes passed in
        pyData1 = self.GetPyData(item1)
        pyData2 = self.GetPyData(item2)
        # For convenience, get the Node Types for the two nodes
        type1 = pyData1.nodetype
        type2 = pyData2.nodetype

        text1 = self.GetItemText(item1)
        text2 = self.GetItemText(item2)

        # If we're inside a Collection Node ...
        if (type1 in ['CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode', 'CollectionNoteNode']) and \
           (type2 in ['CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode', 'CollectionNoteNode']):

            # If Item1 is NOT a Collection and Item 2 IS, or
            #    Item 1 is a Note and Item 2 isn't ...
            if ((type1 != 'CollectionNode') and (type2 == 'CollectionNode')) or \
               ((type1 == 'CollectionNoteNode') and (type2 != 'CollectionNoteNode')):
                # ... return 1 to indicate that the Note comes after the other item
                return 1
            # If Item 1 is a Collection and Item 2 is not, or
            #    Item 2 is a Note and Item 1 isn't ...
            elif ((type1 == 'CollectionNode') and (type2 != 'CollectionNode')) or \
                 ((type1 != 'CollectionNoteNode') and (type2 == 'CollectionNoteNode')):
                # ... return -1 to indicate that the Note comes after the other item
                return -1
            # If both Items are either Collections or Notes
            elif ((type1 == 'CollectionNode') and (type2 == 'CollectionNode')) or \
                 ((type1 == 'CollectionNoteNode') and (type2 == 'CollectionNoteNode')):
                # ... a simple alphabetic sort will do
                if self.GetItemText(item1) < self.GetItemText(item2):
                    return -1
                else:
                    return 1
            # If neither Item is a Collection or a Note ...
            else:
                # ... compare their Sort Order, returning a negative if the order should be switched
                return pyData1.sortOrder - pyData2.sortOrder

        # If we are sorting the Search Quotes, Search Clips and Search Snapshots in a Search Collection ...
        elif (type1 in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']) and (type1 in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']):
            # ... just compare the sort orders
            return pyData1.sortOrder - pyData2.sortOrder


        # If item2 is a Note and item1 is not ...
        elif (type2 in ['LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode',
                        'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']) and \
             (not type1 in ['LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode',
                        'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']):
            return -1
        
        # If item1 is a Note and item2 is not ...
        elif (type1 in ['LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode',
                        'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']) and \
             (not type2 in ['LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode',
                        'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']):
            return 1
        
        # If we have other node types ...
        else:
            # ... a simple alphabetic sort will do
            if self.GetItemText(item1) < self.GetItemText(item2):
                return -1
            else:
                return 1

    def set_image(self, item, icon_name):
        """Set the item's icon image for all states."""
        index = self.icon_list.index(icon_name)
        self.SetItemImage(item, index, wx.TreeItemIcon_Normal)
        self.SetItemImage(item, index, wx.TreeItemIcon_Selected)
        self.SetItemImage(item, index, wx.TreeItemIcon_Expanded)
        self.SetItemImage(item, index, wx.TreeItemIcon_SelectedExpanded)
        
    # FIXME: Doesn't preserve node 'expanded' states
    def refresh_tree(self, evt=None):
        """Load information from database and re-create the tree."""
        self.DeleteAllItems()
        self.create_root_node()
        self.create_series_node()
        self.create_collections_node()
        self.create_kwgroups_node()
        self.create_search_node()
        self.Expand(self.root)
        self.UnselectAll()
        self.SelectItem(self.GetRootItem())

    def OnMotion(self, event):
        """ Detects Mouse Movement in the Database Tree Tab so that we can scroll as needed
            during Drag-and-Drop operations. """

        # Detect if a Drag is currently under way

        # NOTE:  Although I'm sure this code USED to work, it no longer gets called during Dragging.
        #        I don't know if this coincided with going to 2.4.2.4, or whether there is some other aspect
        #        of our Drag-and-drop code that is causing this problem.
        
        if event.Dragging():
            # Get the Mouse position
            (x,y) = event.GetPosition()
            # Get the dimensions of the Database Tree Tab window
            (w, h) = self.GetClientSizeTuple()
            # If we are dragging at the top of the window, scroll down
            if y < 8:
                self.ScrollLines(-2)
            # If we are dragging at the bottom of the window, scroll up
            elif y > h - 8:
                self.ScrollLines(2)
        # If we're not dragging, don't do anything
        else:
            pass
        

    def OnCutCopyBeginDrag(self, event):
        """ Fired on the initiation of a Drag within the Database Tree Tab or by the selection
            of "Cut" or "Copy" from the popup menus.  """
        # The first thing we need to do is determine what tree item is being cut, copied, or dragged.
        # For drag, we just look at the mouse position to determine the item.  However, for Cut and Copy,
        # the mouse position will reflect the position of the the popup menu item selected rather than 
        # the position of the tree item selected.  However, the cutCopyInfo variable will contain that
        # information.

        # SearchCollectionNodes, SearchClipNodes, and SearchSnapshot need their Node Lists included in 
        # the SourceData Object so that the correct tree node can be deleted during "Cut/Move" operations.
        # (I've tried several other approaches, but they've failed.)
        # To be able to do this, we need to initialize the nodeList to empty.
        nodeList = ()

        # Detect if this event has been fired by a "Cut" or "Copy" request.  (Collections, Clips, Snapshots, Keywords,
        # SearchCollections, SearchClips, and SearchSnapshots have
        # "Cut" options as their first menu items and "Copy" as their second menu items.)
        if event.GetId() in [self.cmd_id_start["DocumentNode"],
                             self.cmd_id_start["EpisodeNode"],
                             self.cmd_id_start["CollectionNode"],
                             self.cmd_id_start['QuoteNode'],
                             self.cmd_id_start['ClipNode'],
                             self.cmd_id_start['SnapshotNode'],
                             self.cmd_id_start["NoteNode"],
                             self.cmd_id_start["KeywordNode"],
                             self.cmd_id_start["SearchCollectionNode"],
                             self.cmd_id_start["SearchQuoteNode"],
                             self.cmd_id_start["SearchClipNode"],
                             self.cmd_id_start["SearchSnapshotNode"],
                             self.cmd_id_start["CollectionNode"] + 1,
                             self.cmd_id_start['QuoteNode'] + 1,
                             self.cmd_id_start['ClipNode'] + 1,
                             self.cmd_id_start['SnapshotNode'] + 1,
                             self.cmd_id_start["NoteNode"] + 1,
                             self.cmd_id_start["KeywordNode"] + 1,
                             self.cmd_id_start["SearchCollectionNode"] + 1,
                             self.cmd_id_start["SearchQuoteNode"] + 1,
                             self.cmd_id_start["SearchClipNode"] + 1,
                             self.cmd_id_start["SearchSnapshotNode"] + 1]:
            # If "Cut" or "Copy", get the selected item from cutCopyInfo
            sel_item = self.cutCopyInfo['sourceItem']
            # Reset the DragAndDropObjects' YESTOALL variable, which should no longer indicate that Yes To All has been pressed!
            DragAndDropObjects.YESTOALL = False
            
        # If this method is not triggered by a "Cut" or "Copy" request, it was fired by the initiation of a Drag.
        else:
            # Items in the tree are not automatically selected with a left click.
            # We must select the item that is initially clicked manually!!
            # We do this by looking at the screen point clicked and applying the tree's
            # HitTest method to determine the current item, then actually selecting the item
            if len(self.GetSelections()) <= 1:

                # This line works on Windows, but not on Mac or Linux using wxPython 2.4.1.2  due to a problem with event.GetPoint().
                # pt = event.GetPoint()
                # therfore, this alternate method is used.
                # Get the Mouse Position on the Screen in a more generic way to avoid the problem above
                (windowx, windowy) = wx.GetMousePosition()
                # Translate the Mouse's Screen Position to the Mouse's Control Position
                pt = self.ScreenToClientXY(windowx, windowy)
                # use HitTest to determine the tree item as the screen point indicated.
                sel_item, flags = self.HitTest(pt)
            else:
                sel_item = self.GetSelections()

                # We can work with multiple Quotes, Clips, or Snaps mixed together!
                IsQuoteClipSnap = self.GetPyData(sel_item[0]).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']
                # We can also work with multiple Search Quotes, Search Clips, or Search Snaps mixed together!
                IsSearchQuoteClipSnap = self.GetPyData(sel_item[0]).nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']

##                print "DatabaseTreeTab.OnCutCopyBeginDrag():", IsQuoteClipSnap, self.GetPyData(sel_item[0]).nodetype

                # Check for DRAG and DROP only.  There is a similar section that tests this for right-click menus
                # iterate through the list of selections ...
                for x in sel_item:

##                   print self.GetPyData(x).nodetype, IsQuoteClipSnap, not(self.GetPyData(x).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode'])
                   
                   # ... comparing the node type of the iteration item to the first node type ...
                   if (IsQuoteClipSnap and not (self.GetPyData(x).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode'])) or \
                      (IsSearchQuoteClipSnap and not (self.GetPyData(x).nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode'])) or \
                      (not (IsQuoteClipSnap) and (not (IsSearchQuoteClipSnap)) and (self.GetPyData(sel_item[0]).nodetype != self.GetPyData(x).nodetype)):
                        # ... and if they're different, display an error message ...
                        dlg = Dialogs.ErrorDialog(self.parent, _('All selected items must be the same type to manipulate multiple items at once.'))
                        dlg.ShowModal()
                        dlg.Destroy()
                        # Veto the event
                        event.Veto()
                        # ... then get out of here as we cannot proceed.
                        return

        # Start exception handling to catch Mac problem
        try:
            if not isinstance(sel_item, list):
                # Select the appropriate item in the TreeCtrl
                self.SelectItem(sel_item)
        except wx._core.PyAssertionError:
            pass

        # Initialize the pickled drag-and-drop data object
        pddd = None
        # If sel_item is a LIST, we have multiple tree node selections.
        if isinstance(sel_item, list):
            # Initialize a LIST of drag-and-drop data objects
            dddList = []
            # Iterate through the list of selected nodes
            for item in sel_item:
                # Determine what Item is being cut, copied, or dragged, and grab it's data
                tempNodeName = "%s" % (self.GetItemText(item))
                tempNodeData = self.GetPyData(item)
                
                # If we're dealing with a SearchCollection or SearchClip Node, let's build the nodeList.
                if tempNodeData.nodetype in ['SearchCollectionNode', 'SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode'] :
                    # Start with a Node Pointer
                    tempNode = item
                    # Start the node List with that nodePointer's Text
                    nodeList = (self.GetItemText(tempNode),)
                    # Climb the tree up to the Search Root Node ...
                    while self.GetPyData(tempNode).nodetype != 'SearchRootNode':
                        # Get the Parent Node ...
                        tempNode = self.GetItemParent(tempNode)
                        # And add it's text to the Node List
                        nodeList = (self.GetItemText(tempNode),) + nodeList

                # Create a custom Data Object for Cut and Paste AND Drag and Drop
                ddd = DragAndDropObjects.DataTreeDragDropData(text=tempNodeName, nodetype=tempNodeData.nodetype, nodeList=nodeList, recNum=tempNodeData.recNum, parent=tempNodeData.parent)
                # Add the custom Data Object to the drag-and-drop data items list
                dddList.append(ddd)

            # Use cPickle to convert the data object into a string representation
            pddd = cPickle.dumps(dddList, 1)

        # If you're editing a Keyword Label (to merge keywords) and click on a tree item to end the
        # label edit, the sel_item is NOT OK and you get an error.  Let's make sure we have a good tree
        # item before we go on with the drag.  
        elif sel_item.IsOk():
            # Determine what Item is being cut, copied, or dragged, and grab it's data
            tempNodeName = "%s" % (self.GetItemText(sel_item))
            tempNodeData = self.GetPyData(sel_item)
            
            # If we're dealing with a SearchCollection, SearchClip, or SearchSnapshot Node, let's build the nodeList.
            if tempNodeData.nodetype in ['SearchCollectionNode', 'SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
                # Start with a Node Pointer
                tempNode = sel_item
                # Start the node List with that nodePointer's Text
                nodeList = (self.GetItemText(tempNode),)
                # Climb the tree up to the Search Root Node ...
                while self.GetPyData(tempNode).nodetype != 'SearchRootNode':
                    # Get the Parent Node ...
                    tempNode = self.GetItemParent(tempNode)
                    # And add it's text to the Node List
                    nodeList = (self.GetItemText(tempNode),) + nodeList

            # Create a custom Data Object for Cut and Paste AND Drag and Drop
            ddd = DragAndDropObjects.DataTreeDragDropData(text=tempNodeName, nodetype=tempNodeData.nodetype, nodeList=nodeList, recNum=tempNodeData.recNum, parent=tempNodeData.parent)

            # Use cPickle to convert the data object into a string representation
            pddd = cPickle.dumps(ddd, 1)

        # If a pickled drag-and-drop data object was successfully created ...
        if pddd != None:
            # Now create a wxCustomDataObject for dragging and dropping and
            # assign it a custom Data Format
            cdo = wx.CustomDataObject(wx.CustomDataFormat('DataTreeDragData'))
            # Put the pickled data object in the wxCustomDataObject
            cdo.SetData(pddd)

            # If we have a "Cut" or "Copy" request, we put the pickled CustomDataObject in the Clipboard.
            # If we have a "Drag", we put the pickled CustomDataObject in the DropSource Object.

            # If the event was triggered by a "Cut" or "Copy" request ...
            if event.GetId() in [self.cmd_id_start["DocumentNode"],
                                 self.cmd_id_start["EpisodeNode"],
                                 self.cmd_id_start["CollectionNode"],
                                 self.cmd_id_start['QuoteNode'],
                                 self.cmd_id_start['ClipNode'],
                                 self.cmd_id_start['SnapshotNode'],
                                 self.cmd_id_start["NoteNode"],
                                 self.cmd_id_start["KeywordNode"],
                                 self.cmd_id_start["SearchCollectionNode"],
                                 self.cmd_id_start["SearchQuoteNode"],
                                 self.cmd_id_start["SearchClipNode"],
                                 self.cmd_id_start["SearchSnapshotNode"],
                                 self.cmd_id_start["CollectionNode"] + 1,
                                 self.cmd_id_start['QuoteNode'] + 1,
                                 self.cmd_id_start['ClipNode'] + 1,
                                 self.cmd_id_start['SnapshotNode'] + 1,
                                 self.cmd_id_start["NoteNode"] + 1,
                                 self.cmd_id_start["KeywordNode"] + 1,
                                 self.cmd_id_start["SearchCollectionNode"] + 1,
                                 self.cmd_id_start["SearchQuoteNode"] + 1,
                                 self.cmd_id_start["SearchClipNode"] + 1,
                                 self.cmd_id_start["SearchSnapshotNode"] + 1]:
                # Open the Clipboard
                wx.TheClipboard.Open()
                # ... put the data in the clipboard ...
                wx.TheClipboard.SetData(cdo)
                # Close the Clipboard
                wx.TheClipboard.Close()
                
            # If the event was triggered by a "Drag" request ...
            else:
                # Create a Custom DropSource Object.  The custom drop source object
                # accepts the tree as a parameter so it can query it about where the
                # drop is supposed to be occurring to see if it will allow the drop.
                tds = DragAndDropObjects.DataTreeDropSource(self)
                # Associate the Data with the Drop Source Object
                tds.SetData(cdo)
                # Initiate the Drag Operation
                dragResult = tds.DoDragDrop(True)
                # We do the actual processing in the DropTarget object, as we have access to the Dragged Data and the
                # Drop Target's information there but not here.

                # Because the DropSource GiveFeedback Method can change the cursor, I find that I
                # need to reset it to "normal" here or it can get stuck as a "No_Entry" cursor if
                # a Drop is abandoned.
                self.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

    def create_root_node(self):
        """ Build the entire Database Tree """
        # Include the Database Name as part of the Database Root
        prompt = _('Database: %s')
        # Note:  French was having trouble with UTF-8 prompts without the following line.
        #        We use 'utf8' rather than TransanaGlobal.encoding because it's NOT the database's encoding
        #        but the prompt encoding (always utf-8) that we need here!
        if ('unicode' in wx.PlatformInfo) and (type(prompt).__name__ == 'str'):
            prompt = unicode(prompt, 'utf8')  # TransanaGlobal.encoding)
        prompt = prompt % TransanaGlobal.configData.database
        # Add the Tree's root node
        self.root = self.AddRoot(prompt)
        # Add the node's image and node data
        if TransanaConstants.singleUserVersion:
            self.set_image(self.root, "db")
        elif TransanaGlobal.configData.ssl:
            self.set_image(self.root, "db_locked")
        else:
            self.set_image(self.root, "db_unlocked")
        nodedata = _NodeData(nodetype='Root')                    # Identify this as the Root node
        self.SetPyData(self.root, nodedata)                      # Associate this data with the node
      
    def create_series_node(self):
        """ Create the Library node and populate it with all appropriate data """
        # We need to keep track of the nodes so we can add sub-nodes quickly.  A dictionary of dictionaries will do this well. 
        mapDict = {'Libraries' : {}, 'Episode' : {}, 'Transcript' : {}, 'Document' : {}}                 
        # Add the root 'Libraries' node
        root_item = self.AppendItem(self.root, _('Libraries'))
        # Add the node's image and node data
        nodedata = _NodeData(nodetype='LibraryRootNode')         # Idenfify this as the Library Root node
        self.SetPyData(root_item, nodedata)                      # Associate this data with the node
        self.set_image(root_item, "LibraryRoot16")

        # The following code is RADICALLY faster than the original version, like 1000% faster.  It accomplishes this by
        # minimizing the number of database calls and tracking the tree nodes with a map dictionary so that we can
        # easily locate the node we want to add a child node to.

        # Populate the tree with all Library records
        for (libraryNo, libraryID) in DBInterface.list_of_series():
            # Create the tree node
            item = self.AppendItem(root_item, libraryID)
            # Add the node's image and node data
            nodedata = _NodeData(nodetype='LibraryNode', recNum=libraryNo)          # Identify this as a Library node
            self.SetPyData(item, nodedata)                       # Associate this data with the node
            self.set_image(item, "Library16")
            # Add the new node to the map dictionary
            mapDict['Libraries'][libraryNo] = item

        # Populate the tree with all Documents AND Episodes
        tmpDict = DBInterface.dictionary_of_documents_and_episodes()
        keys = tmpDict.keys()
        keys.sort()
        for key in keys:
            (objType, objNum, objParentNum) = tmpDict[key]
            # Check to see if the Clip's parent collection is in the tree.  It should be there, but I did
            # have a testing database where one collection was missing, despite the presence of Clips and Notes.
            if mapDict['Libraries'].has_key(objParentNum):
                # Exclude Documents in Transana Standard
                if TransanaConstants.proVersion or objType != 'Document':
                    # Find the correct Library node using the map dictionary
                    libitem = mapDict['Libraries'][objParentNum]
                    # Create the tree node
                    deitem = self.AppendItem(libitem, key[0])
                    if objType == 'Document':
                        # Add the node's image and node data
                        nodedata = _NodeData(nodetype='DocumentNode', recNum=objNum, parent=objParentNum)  # Identify this as a Document node
                        self.set_image(deitem, "Document16")
                        # Add the new node to the map dictionary
                        mapDict['Document'][objNum] = deitem
                    elif objType == 'Episode':
                        # Add the node's image and node data
                        nodedata = _NodeData(nodetype='EpisodeNode', recNum=objNum, parent=objParentNum)     # Identify this as an Episode node
                        self.set_image(deitem, "Episode16")
                        # Add the new node to the map dictionary
                        mapDict['Episode'][objNum] = deitem
                    self.SetPyData(deitem, nodedata)                  # Associate this data with the node
            # This shouldn't happen to anyone but me.  God, I hope not, anyway.
            else:
                print "ABANDONED %s RECORD!" % objType.upper(), objNum, objParentNum

        # Populate the tree with all Episode Transcripts
        for (transcriptNo, transcriptID, transcriptEpisodeNo) in DBInterface.list_of_episode_transcripts():
            # Find the correct Library node using the map dictionary
            epitem = mapDict['Episode'][transcriptEpisodeNo]
            # Create the tree node
            titem = self.AppendItem(epitem, transcriptID)
            # Add the node's image and node data
            nodedata = _NodeData(nodetype='TranscriptNode', recNum=transcriptNo, parent=transcriptEpisodeNo)  # Identify this as a Transcript node
            self.SetPyData(titem, nodedata)                  # Associate this data with the node
            self.set_image(titem, "Transcript16")
            # Add the new node to the map dictionary
            mapDict['Transcript'][transcriptNo] = titem

        # Now add all the Notes to the objects in the Library node of the database tree
        for (noteNum, noteID, libraryNum, episodeNum, transcriptNum, collectNum, clipNum, snapshotNum, documentNum, quoteNum) in \
            DBInterface.list_of_node_notes(LibraryNode=True):
            # Find the correct Library, Episode, or Transcript node using the map dictionary
            if libraryNum > 0:
                item = mapDict['Libraries'][libraryNum]
                noteNodeType = 'LibraryNoteNode'
            elif episodeNum > 0:
                item = mapDict['Episode'][episodeNum]
                noteNodeType = 'EpisodeNoteNode'
            elif transcriptNum > 0:
                item = mapDict['Transcript'][transcriptNum]
                noteNodeType = 'TranscriptNoteNode'
            elif TransanaConstants.proVersion and documentNum > 0:
                item = mapDict['Document'][documentNum]
                noteNodeType = 'DocumentNoteNode'
##            elif quoteNum > 0:
##                item = mapDict['Quote'][documentNum]
##                noteNodeType = 'QuoteNoteNode'
            # Create the tree node
            noteitem = self.AppendItem(item, noteID)
            # Add the node's image and node data
            nodedata = _NodeData(nodetype=noteNodeType, recNum=noteNum)  # Identify this as a Note node
            self.SetPyData(noteitem, nodedata)                  # Associate this data with the node
            self.set_image(noteitem, "Note16")

    def create_collections_node(self):
        """ Create the Collections node and populate it with all appropriate data """
        # We need to keep track of the nodes so we can add sub-nodes quickly.  A dictionary of dictionaries will do this well. 
        mapDict = {'Collection' : {}, 'Clip' : {}, 'Snapshot' : {}, 'Quote' : {}}
        # Because of the way data is returned from the database, we will occasionally run into a nested collection
        # whose parent has not yet been added to the tree.  We need a place to store these records for later processing
        # after their parent has been added.
        deferredItems = []

        # Add the root 'Collections' node
        root_item = self.AppendItem(self.root, _("Collections"))
        # Add the node's image and node data
        nodedata = _NodeData(nodetype='CollectionsRootNode')     # Identify this as the Collections Root node
        self.SetPyData(root_item, nodedata)                      # Associate this data with the node
        self.set_image(root_item, "Collection16")
        # Because of the nature of nested collections, we should put the Collections Root into the map dictionary
        # so that first-level nodes can find it as their parent
        mapDict['Collection'][0] = root_item

        # The following code is RADICALLY faster than the original version, like 1000% faster.  It accomplishes this by
        # minimizing the number of database calls and tracking the tree nodes with a map dictionary so that we can
        # easily locate the node we want to add a child node to.

        # Populate the tree with all Collection records
        for (collNo, collID, parentCollNo) in DBInterface.list_of_all_collections():

            if DEBUG:
                print "Collections:", collNo, collID.encode('utf8'), parentCollNo
            
            # First, let's see if the parent collection is in the map dictionary.
            if mapDict['Collection'].has_key(parentCollNo):
                # If so, we can identify the parent collection node (including the root node for parentless Collections)
                # using the map dictionary
                parentItem = mapDict['Collection'][parentCollNo]
                # Create the tree node
                item = self.AppendItem(parentItem, collID)
                # Identify this as a Collection node with the proper Node Data
                nodedata = _NodeData(nodetype='CollectionNode', recNum=collNo, parent=parentCollNo)
                # Associate this data with the node
                self.SetPyData(item, nodedata)
                # Select the proper image
                self.set_image(item, "Collection16")
                # Add the new node to the map dictionary
                mapDict['Collection'][collNo] = item

                # We need to check the items waiting to be processed to see if we've just added the parent
                # collection for any of the items in the list.  If the list is empty, though, we don't need to bother.
                placementMade = (len(deferredItems) > 0)
                # We do this in a while loop, as each item from the list that gets added to the tree could be the
                # parent of other items in the list.
                while placementMade:
                    # Re-initialize the while loop variable, assuming that no items will be found
                    placementMade = False
                    # Now see if any of the deferred items can be added!  Loop through the list ...
                    # We can't just use a for loop, as we delete items from the list as we go!
                    # So start by defining the list index and the number of items in the list
                    index = 0
                    endPoint = len(deferredItems)
                    # As long as the index is less than the last list item ...
                    while index < endPoint:  # for index in range(len(deferredItems)):
                        # Get the data from the list item
                        (dCollNo, dCollID, dParentCollNo) = deferredItems[index]

                        if DEBUG:
                            print "Deferred Collections:", index, endPoint, dCollNo, dCollID, dParentCollNo
                            
                        # See if the parent collection has now been added to the tree and to the map dictionary
                        if mapDict['Collection'].has_key(dParentCollNo):
                            # We can identify the parent node using the map dictionary
                            parentItem = mapDict['Collection'][dParentCollNo]
                            # Create the tree node
                            item = self.AppendItem(parentItem, dCollID)
                            # Identify this as a Collection node with the proper Node Data
                            nodedata = _NodeData(nodetype='CollectionNode', recNum=dCollNo, parent=dParentCollNo)
                            # Associate this data with the node
                            self.SetPyData(item, nodedata)
                            # Select the proper image
                            self.set_image(item, "Collection16")
                            # Add the new node to the map dictionary
                            mapDict['Collection'][dCollNo] = item
                            # We need to indicate to the while loop that we found an entry that could be the parent of other entries
                            placementMade = True
                            # We need to remove the item we just added to the tree from the deferred items list
                            del deferredItems[index]
                            # If we delete the item from the list, we reduce the End Point by one and DO NOT increase the index
                            endPoint -= 1
                        # If the parent collection is NOT in the mapping dictionalry yet ...
                        else:
                            # ... just move on to the next item.  We can't do anything with this yet.
                            index += 1
                    
            # If the Collection's parent is not yet in the database tree or the Map dictionary ...
            else:
                # ... we need to place that collection in the list of items to process later, once the parent Collection
                # has been added to the database tree
                deferredItems.append((collNo, collID, parentCollNo))
                
        # Populate the tree with all Clip records
        for (clipNo, clipID, collNo, sourceNo, sortOrder) in DBInterface.list_of_clips():
            # Check to see if the Clip's parent collection is in the tree.  It should be there, but I did
            # have a testing database where one collection was missing, despite the presence of Clips and Notes.
            if mapDict['Collection'].has_key(collNo):
                # First, let's see if the parent collection is in the map dictionary.
                item = mapDict['Collection'][collNo]
                # Create the tree node
                clip_item = self.AppendItem(item, clipID)
                # Create the node data and assign the node's image
                nodedata = _NodeData(nodetype='ClipNode', recNum=clipNo, parent=collNo, sortOrder=sortOrder, sourceObj=sourceNo)       # Identify this as a Clip node
                self.SetPyData(clip_item, nodedata)                           # Associate this data with the node
                self.set_image(clip_item, "Clip16")
                # Add the new node to the map dictionary
                mapDict['Clip'][clipNo] = clip_item
            # This shouldn't happen to anyone but me.  God, I hope not, anyway.
            else:
                print "ABANDONED CLIP RECORD!" , clipNo, clipID.encode('utf8'), collNo

        # If we're in a Pro version, not the Standard Version ...
        if TransanaConstants.proVersion:
            # Populate the tree with all Quote records
            for (quoteNum, quoteID, collNum, sourceDoc, sortOrder) in DBInterface.list_of_quotes():
                # Check to see if the Quote's parent collection is in the tree.  It should be there.
                if mapDict['Collection'].has_key(collNum):
                    # First, let's see if the parent collection is in the map dictionary.
                    item = mapDict['Collection'][collNum]
                    # Create the tree node
                    quote_item = self.AppendItem(item, quoteID)
                    # Create the node data and assign the node's image
                    nodedata = _NodeData(nodetype='QuoteNode', recNum=quoteNum, parent=collNum, sortOrder=sortOrder, sourceObj=sourceDoc)
                    self.SetPyData(quote_item, nodedata)                           # Associate this data with the node
                    self.set_image(quote_item, "Quote16")
                    # Add the new node to the map dictionary
                    mapDict['Quote'][quoteNum] = quote_item
                # This shouldn't happen to anyone but me.  God, I hope not, anyway.
                else:
                    print "ABANDONED QUOTE RECORD!" , quoteNum, quoteID.encode('utf8'), collNum

            # Populate the tree with all Snapshot records
            for (snapshotNo, snapshotID, collNo, sortOrder) in DBInterface.list_of_snapshots():
                # Check to see if the Snapshot's parent collection is in the tree.  It should be there, but I did
                # have a testing database where one collection was missing, despite the presence of Clips and Notes.
                if mapDict['Collection'].has_key(collNo):
                    # First, let's see if the parent collection is in the map dictionary.
                    item = mapDict['Collection'][collNo]
                    # Create the tree node
                    snapshot_item = self.AppendItem(item, snapshotID)
                    # Create the node data and assign the node's image
                    nodedata = _NodeData(nodetype='SnapshotNode', recNum=snapshotNo, parent=collNo, sortOrder=sortOrder)       # Identify this as a Snapshot node
                    self.SetPyData(snapshot_item, nodedata)                           # Associate this data with the node
                    self.set_image(snapshot_item, "Snapshot16")
                    # Add the new node to the map dictionary
                    mapDict['Snapshot'][snapshotNo] = snapshot_item
                # This shouldn't happen to anyone but me.  God, I hope not, anyway.
                else:
                    print "ABANDONED SNAPSHOT RECORD!" , snapshotNo, snapshotID.encode('utf8'), collNo

        # For each Collection ...
        for key in mapDict['Collection'].keys():
            # ... sort the collection's children!
            self.SortChildren(mapDict['Collection'][key])

        # Now add all the Notes to the objects in the Collection node of the database tree
        for (noteNum, noteID, libraryNum, episodeNum, transcriptNum, collectNum, clipNum, snapshotNum, documentNum, quoteNum) in \
            DBInterface.list_of_node_notes(CollectionNode=True):
            item = None
            # Find the correct Collection or Clip node using the map dictionary
            if collectNum > 0:
                if mapDict['Collection'].has_key(collectNum):
                    item = mapDict['Collection'][collectNum]
                    noteNodeType = 'CollectionNoteNode'
                else:
                    print "ABANDONED COLLECTION NOTE RECORD!", noteNum, noteID.encode('utf8'), collectNum
            elif clipNum > 0:
                if mapDict['Clip'].has_key(clipNum):
                    item = mapDict['Clip'][clipNum]
                    noteNodeType = 'ClipNoteNode'
                else:
                    print "ABANDONED CLIP NOTE RECORD!", noteNum, noteID.encode('utf8'), clipNum
            elif (snapshotNum > 0) and TransanaConstants.proVersion:
                if mapDict['Snapshot'].has_key(snapshotNum):
                    item = mapDict['Snapshot'][snapshotNum]
                    noteNodeType = 'SnapshotNoteNode'
                else:
                    print "ABANDONED SNAPSHOT NOTE RECORD!", noteNum, noteID.encode('utf8'), snapshotNum
##            elif (documentNum > 0) and TransanaConstants.proVersion:
##                if mapDict['Document'].has_key(documentNum):
##                    item = mapDict['Document'][documentNum]
##                    noteNodeType = 'DocumentNoteNode'
##                else:
##                    print "ABANDONED DOCUMENT NOTE RECORD!", noteNum, noteID.encode('utf8'), documentNum
            elif (quoteNum > 0) and TransanaConstants.proVersion:
                if mapDict['Quote'].has_key(quoteNum):
                    item = mapDict['Quote'][quoteNum]
                    noteNodeType = 'QuoteNoteNode'
                else:
                    print "ABANDONED QUOTE NOTE RECORD!", noteNum, noteID.encode('utf8'), quoteNum
            if item != None:
                # Create the tree node
                noteitem = self.AppendItem(item, noteID)
                # Add the node's image and node data
                nodedata = _NodeData(nodetype=noteNodeType, recNum=noteNum)  # Identify this as a Note node
                self.SetPyData(noteitem, nodedata)                  # Associate this data with the node
                self.set_image(noteitem, "Note16")

    def create_kwgroups_node(self):
        """ Create the Keywords node and populate it with all appropriate data """
        # Add the root 'Keywords' node
        kwg_root = self.AppendItem(self.root, _("Keywords"))
        # Add the node's image and node data
        nodedata = _NodeData(nodetype='KeywordRootNode')         # Identify this as the Keywords Root node
        self.SetPyData(kwg_root, nodedata)                       # Associate this data with the node
        self.set_image(kwg_root, "KeywordRoot16")
        # Since there are times when we need to refresh the Keywords node but not the whole tree, a separate
        # method has been created for populating an existing Keywords Node.  We can just call it.
        self.refresh_kwgroups_node()
        
    def refresh_kwgroups_node(self):
        """ Refresh the Keywords Node of the Database Tree """
        selItems = self.GetSelections()
        # Initialize keyword groups to an empty list.  This list keeps track of the defined Keyword Groups for use elsewhere
        self.kwgroups = []
        # We need to keep track of the nodes so we can add sub-nodes quickly.  A dictionary of dictionaries will do this well. 
        mapDict = {}
        # We need to locate the existing Kewords Root Node
        kwg_root = self.select_Node((_("Keywords"),), 'KeywordRootNode')
        # Now, we can clear out the Keywords Node completely to start over.
        self.DeleteChildren(kwg_root)
        # Let's add the root node to our Keyword List
        self.kwgroups.append(kwg_root)
        # Get all Keyword Group : Keyword pairs from the database
        for (kwg, kw) in DBInterface.list_of_all_keywords():
            # Check to see if the Keyword Group has already been added to the Database Tree and the Map dictionary 
            if not mapDict.has_key(kwg.upper()):
                # If not, add the Keyword Group to the Tree
                kwg_item = self.AppendItem(kwg_root, kwg)
                # Specify the Keyword Group's node data and image
                nodedata = _NodeData(nodetype='KeywordGroupNode')    # Identify this as a Keyword Group node
                self.SetPyData(kwg_item, nodedata)                   # Associate this data with the node
                self.set_image(kwg_item, "KeywordGroup16")
                # Add the Keyword Group to our Keyword Groups List
                self.kwgroups.append(kwg_item)
                # Add the Keyword Group to our map dictionary, and note the associated tree node as "item"
                mapDict[kwg.upper()] = {'item' : kwg_item}
            # If the Keyword Group IS in the tree and map already ...
            else:
                # ... we can identify the corresponding tree node from the map dictionary
                kwg_item = mapDict[kwg.upper()]['item']
            # Add the Keyword to the database tree
            kw_item = self.AppendItem(kwg_item, kw)
            # Specify the Keyword's node data and image
            nodedata = _NodeData(nodetype='KeywordNode', parent=kwg)     # Identify this as a Keyword node
            self.SetPyData(kw_item, nodedata)                # Associate this data with the node
            self.set_image(kw_item, "Keyword16")
            # Add the Keyword to the map dictionary, pointing to the keyword's tree node
            mapDict[kwg.upper()][kw] = kw_item

        # Get all Keyword Examples from the database
        keywordExamples = DBInterface.list_of_keyword_examples()

        # NOTE:  This would be more efficient if the DBInterface.list_of_keyword_examples() method passed all necessary
        #        information from the database rather than requiring that we load each Clip to determine its ID and parent
        #        Collection.  However, I suspect that Keyword Examples are rare enough that it's not a major issue.

        # Iterate through the examples
        for (episodeNum, clipNum, snapshotNum, kwg, kw, example) in keywordExamples:
            # Load the indicated clip.  We can speed the load by not loading the Clip Transcript(s)
            exampleClip = Clip.Clip(clipNum, skipText=True)
            # Determine where it should be displayed in the Node Structure.
            # (Keyword Root, Keyword Group, Keyword, Example Clip Name)
            nodeData = (_('Keywords'), kwg, kw, exampleClip.id)
            # Add the Keyword Example Node to the Database Tree Tab, but don't expand the nodes
            self.add_Node("KeywordExampleNode", nodeData, exampleClip.number, exampleClip.collection_num, expandNode=False)

        # Reset the selection in the Tree to what it was before we called this method
        if len(selItems) > 0:
            self.UnselectAll()
            self.SelectItem(selItems[0])
        # Refresh the Tree Node so that changes are displayed appropriately (such as new children are indicated if the node was empty)
        self.Refresh()

    def add_note_nodes(self, note_ids, item, **parent_num):
        """ Add the notes specified in note_ids to item """
        if len(note_ids) > 0:
            if parent_num.has_key('Libraries'):
                noteNodeType = 'LibraryNoteNode'
            elif parent_num.has_key('Document'):
                noteNodeType = 'DocumentNoteNode'
            elif parent_num.has_key('Episode'):
                noteNodeType = 'EpisodeNoteNode'
            elif parent_num.has_key('Transcript'):
                noteNodeType = 'TranscriptNoteNode'
            elif parent_num.has_key('Collection'):
                noteNodeType = 'CollectionNoteNode'
            elif parent_num.has_key('Quote'):
                noteNodeType = 'QuoteNoteNode'
            elif parent_num.has_key('Clip'):
                noteNodeType = 'ClipNoteNode'
            elif parent_num.has_key('Snapshot'):
                noteNodeType = 'SnapshotNoteNode'
            else:
                noteNodeType = 'NoteNode'
            for n in note_ids:
                noteitem = self.AppendItem(item, n)
                self.set_image(noteitem, "Note16")
                note = Note.Note(n, **parent_num)
                nodedata = _NodeData(nodetype=noteNodeType, recNum=note.number)  # Identify this as a Note node
                self.SetPyData(noteitem, nodedata)                          # Associate this data with the node
                del note

    def create_search_node(self):
        """ Create the root Search node """
        self.searches = []
        # The "Search" node itself is always item 0 in the node list
        search_root = self.AppendItem(self.root, _("Search"))
        nodedata = _NodeData(nodetype='SearchRootNode')          # Identify this as the Search Root node
        self.SetPyData(search_root, nodedata)                    # Associate this data with the node
        self.set_image(search_root, "SearchRoot16")
        self.searches.append(search_root)

    def UpdateExpectedNodeType(self, expectedNodeType, nodeListPos, nodeData, nodeType):
        """ This function returns the node type of the Next Node that should be examined.
            Used in crawling the Database Tree.
            Paramters are:
              expectedNodeType  The current anticipated Node Type
              nodeListPos       The position in the current Node List of the current node
              nodeData          The current Node List
              nodeType          The Node Type of the LAST element in the Node List """

        # First, let's see if we're dealing with a NOTE, as the next node is different if we are.
        if ((nodeListPos == len(nodeData) - 1) and (nodeType in ['NoteNode', 'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode',
                                                                 'CollectionNoteNode', 'ClipNoteNode', 'SnapshotNoteNode', 'DocumentNoteNode',
                                                                 'QuoteNoteNode'])):
            expectedNodeType = nodeType
        # For DocumentNotes only, we need to move from Library to Document at the second-to-last node
        elif ((nodeListPos == len(nodeData) - 2) and (nodeType == 'DocumentNoteNode')):
            expectedNodeType = 'DocumentNode'
        # For EpisodeNotes only, we need to move from Library to episode at the second-to-last node
        elif ((nodeListPos == len(nodeData) - 2) and (nodeType == 'EpisodeNoteNode')):
            expectedNodeType = 'EpisodeNode'
        # For TranscriptNotes only, we need to move from Library to Episode at the third-to-last node
        elif ((nodeListPos == len(nodeData) - 3) and (nodeType == 'TranscriptNoteNode')):
            expectedNodeType = 'EpisodeNode'
        # For QuoteNotes only, we need to move from Collection to Quote at the second-to-last node
        elif ((nodeListPos == len(nodeData) - 2) and (nodeType == 'QuoteNoteNode')):
            expectedNodeType = 'QuoteNode'
        # For ClipNotes only, we need to move from Collection to Clip at the second-to-last node
        elif ((nodeListPos == len(nodeData) - 2) and (nodeType == 'ClipNoteNode')):
            expectedNodeType = 'ClipNode'
        # For SnapshotNotes only, we need to move from Collection to Snapshot at the second-to-last node
        elif ((nodeListPos == len(nodeData) - 2) and (nodeType == 'SnapshotNoteNode')):
            expectedNodeType = 'SnapshotNode'
        elif expectedNodeType == 'LibraryRootNode':
            expectedNodeType = 'LibraryNode'
##        elif expectedNodeType == 'LibraryNode':
##            expectedNodeType = 'EpisodeNode'
        # LibraryNode only advances to DocumentNode if we're reaching the end of the nodeList AND we're supposed to move on to a Document
        elif (expectedNodeType == 'LibraryNode') and ((nodeListPos == len(nodeData) - 1) and (nodeType in ['DocumentNode'])):
            expectedNodeType = 'DocumentNode'
        # LibraryNode only advances to EpisodeNode if we're nearing the end of the nodeList AND we're supposed to move on to an Episode or a Transcript
        elif (expectedNodeType == 'LibraryNode') and ((nodeListPos >= len(nodeData) - 2) and (nodeType in ['EpisodeNode', 'TranscriptNode'])):
            expectedNodeType = 'EpisodeNode'
        elif expectedNodeType == 'EpisodeNode':
            expectedNodeType = 'TranscriptNode'
        elif expectedNodeType == 'CollectionsRootNode':
            expectedNodeType = 'CollectionNode'
        # CollectionNode only advances to QuoteNode if we're reaching the end of the nodeList AND we're supposed to move on to a Quote
        elif (expectedNodeType == 'CollectionNode') and ((nodeListPos == len(nodeData) - 1) and (nodeType in ['QuoteNode'])):
            expectedNodeType = 'QuoteNode'
        # CollectionNode only advances to ClipNode if we're reaching the end of the nodeList AND we're supposed to move on to a Clip
        elif (expectedNodeType == 'CollectionNode') and ((nodeListPos == len(nodeData) - 1) and (nodeType in ['ClipNode'])):
            expectedNodeType = 'ClipNode'
        # CollectionNode only advances to SnapshotNode if we're reaching the end of the nodeList AND we're supposed to move on to a Snapshot
        elif (expectedNodeType == 'CollectionNode') and ((nodeListPos == len(nodeData) - 1) and (nodeType in ['SnapshotNode'])):
            expectedNodeType = 'SnapshotNode'
        elif expectedNodeType == 'KeywordRootNode':
            expectedNodeType = 'KeywordGroupNode'
        elif expectedNodeType == 'KeywordGroupNode':
            expectedNodeType = 'KeywordNode'
        elif expectedNodeType == 'KeywordNode':
            expectedNodeType = 'KeywordExampleNode'
        elif expectedNodeType == 'SearchRootNode':
            expectedNodeType = 'SearchResultsNode'
        elif (expectedNodeType == 'SearchResultsNode') and (nodeType in ['SearchLibraryNode', 'SearchEpisodeNode', 'SearchTranscriptNode', \
                                                                         'SearchDocumentNode']):
            expectedNodeType = 'SearchLibraryNode'
        elif expectedNodeType == 'SearchLibraryNode':
            if nodeType in ['SearchDocumentNode']:
                expectedNodeType = 'SearchDocumentNode'
            elif nodeType in ['SearchEpisodeNode', 'SearchTranscriptNode']:
                expectedNodeType = 'SearchEpisodeNode'
        elif expectedNodeType == 'SearchEpisodeNode':
            expectedNodeType = 'SearchTranscriptNode'
        elif (expectedNodeType == 'SearchResultsNode') and (nodeType in ['SearchCollectionNode', 'SearchQuoteNode', 'SearchClipNode', \
                                                                         'SearchSnapshotNode']):
            expectedNodeType = 'SearchCollectionNode'
        elif (expectedNodeType == 'SearchCollectionNode') and ((nodeListPos == len(nodeData) - 1) and (nodeType in ['SearchQuoteNode'])):
            expectedNodeType = 'SearchQuoteNode'
        elif (expectedNodeType == 'SearchCollectionNode') and ((nodeListPos == len(nodeData) - 1) and (nodeType in ['SearchClipNode'])):
            expectedNodeType = 'SearchClipNode'
        elif (expectedNodeType == 'SearchCollectionNode') and ((nodeListPos == len(nodeData) - 1) and (nodeType in ['SearchSnapshotNode'])):
            expectedNodeType = 'SearchSnapshotNode'
        elif expectedNodeType == 'Node':
            expectedNodeType = 'Node'
        return expectedNodeType

    def Evaluate(self, node, nodeType, child, childData):
        """ The logic for traversing tree nodes gets complicated.  This boolean function encapsulates the decision logic. """
        allNoteNodeTypes = ['NoteNode', 'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode', 'ClipNoteNode',
                            'SnapshotNoteNode', 'DocumentNoteNode', 'QuoteNodeNode']

        # We continue moving down the list of nodes if...
        #   ... we are not yet at the end of the list AND ...
        # ((we're not past our alpha place (UPPER() calls fix case issues!) and (nodetypes are the same or (both are at least Notes)) or
        #  (we're dealing with a Note and we're not to the notes yet)) AND ...
        # (we don't have a Collection or we haven't started looking at Clips, Quotes, Snapshots ) or
        # we've got a SearchCollection to position after the SearchLibrary nodes
        result = child.IsOk() and \
                 (((node.upper() > self.GetItemText(child).upper()) and \
                   (((nodeType == childData.nodetype)) or \
                    ((nodeType in allNoteNodeTypes) and (childData.nodetype in allNoteNodeTypes)))) or \
                  ((nodeType in allNoteNodeTypes) and not(childData.nodetype in allNoteNodeTypes))) or \
                 ((nodeType in ['QuoteNode', 'ClipNode', 'SnapshotNode']) and \
                  (childData.nodetype in ['CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode'])) or \
                 ((nodeType == 'SearchCollectionNode') and (childData.nodetype == 'SearchLibraryNode')) or \
                 ((nodeType in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']) and \
                  (childData.nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']))
        return result
        
    def add_Node(self, nodeType, nodeData, nodeRecNum, nodeParent, sortOrder=None, expandNode = True, insertPos = None,
                 avoidRecursiveYields = False):
        """ This method is used to add nodes to the tree after it has been built.
            nodeType is the type of node to be added, and nodeData is a list that gives the tree structure
            that describes where the node should be added. """

        # ERROR CHECKING

        # This checks for an "expandNode" that doesn't have the leading parameter name (when sortOrder was added)
        if isinstance(sortOrder, bool):
            import TransanaExceptions
            raise TransanaExceptions.NotImplementedError

        # This checks for a Quote, Clip or Snapshot without a Sort Order
        if nodeType in [ 'QuoteNode', 'ClipNode', 'SnapshotNode' ]:
            if sortOrder == None:
                import TransanaExceptions
                raise TransanaExceptions.NotImplementedError

        # Start at the tree's Root Node
        currentNode = self.GetRootItem()
        # See if we need to encode the first element of the tuple/list
        # (This shouldn't have to be so complicated!  But it's different on Mac-PPC and in different circumstances!)
        if type(_(nodeData[0])).__name__ == 'str':
            # We need to translate the root node name and convert it to Unicode.
            if type(nodeData) == type(()):
                nodeData = (unicode(_(nodeData[0]), 'utf8'),) + nodeData[1:]
            elif type(nodeData) == type([]):
                nodeData = [unicode(_(nodeData[0]), 'utf8')] + nodeData[1:]
        # Even if we don't need to encode it, we need to translate it
        else:
            # We need to translate the root node name and convert it to Unicode.
            if type(nodeData) == type(()):
                nodeData = (_(nodeData[0]),) + nodeData[1:]
            elif type(nodeData) == type([]):
                nodeData = [_(nodeData[0])] + nodeData[1:]

        if DEBUG:
            print "DatabaseTreeTab.add_Node():"
            print "Root node = %s" % self.GetItemText(currentNode)
            print "nodeData =", nodeData
            print 'nodeType =', nodeType
            print 'sortOrder =', sortOrder
            print 'insertPos =', insertPos

        # Having nodes and subnodes with the same name causes a variety of problems.  We need to track how far
        # down the tree branches we are to keep track of what object NodeTypes we should be working with.
        nodeListPos = 0
        if nodeType in ['LibraryRootNode', 'LibraryNode', 'EpisodeNode', 'TranscriptNode', 'DocumentNode',
                        'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'DocumentNoteNode']:
            expectedNodeType = 'LibraryRootNode'
        elif nodeType in ['CollectionsRootNode', 'CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode',
                          'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']:
            expectedNodeType = 'CollectionsRootNode'
        elif nodeType in ['KeywordRootNode', 'KeywordGroupNode', 'KeywordNode', 'KeywordExampleNode']:
            expectedNodeType = 'KeywordRootNode'
        elif nodeType in ['SearchRootNode', 'SearchResultsNode', 'SearchLibraryNode', 'SearchDocumentNode', 'SearchEpisodeNode',
                          'SearchTranscriptNode', 'SearchCollectionNode', 'SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
            expectedNodeType = 'SearchRootNode'

        if DEBUG:
            print 'expectedNodeType =', expectedNodeType
            print 'nodeData = ', nodeData


        indexPos = 0

        for node in nodeData:

            node = Misc.unistrip(node)

            if DEBUG:
                print "Looking for '%s'  (%s)" % (node, type(node))
                print "Current nodeListPos = %d, expectedNodeType = %s" % (nodeListPos, expectedNodeType)

            notDone = True

            if DEBUG:
                print "Getting children for ", self.GetItemText(currentNode)
                
            (childNode, cookieItem) = self.GetFirstChild(currentNode)

            if DEBUG:
                print "GetFirstChild call complete"

            while notDone:

		if childNode.IsOk():

                    if DEBUG:
                        print "childNode is OK"
                        
                    itemText = Misc.unistrip(self.GetItemText(childNode))
                    
                    if DEBUG:
                        print u"Looking in %s " % self.GetItemText(currentNode)
#                        print u"at %s " % self.GetItemText(childNode)
                        print u"for %s." % node

                    # Let's get the child Node's Data
                    childNodeData = self.GetPyData(childNode)
                else:

                    if DEBUG:
                        print "childNode is NOT OK"
                        
                    itemText = ''
                    childNodeData = None
                    
                if DEBUG:
                    print nodeType, childNode.IsOk()

                # To accept a node and climb on, the node text must match the text of the node being sought.
                # If the node being added is a Keyword Example Node, we also need to make sure that the node being
                # sought's Record Number matches the node being added's Record Number.  This is because it IS possible
                # to have two clips with the same name in different collections applied as examples of the same Keyword.

                if DEBUG:
                    print 'NodeTypes:', nodeType, 
                    if childNode.IsOk():
                        print childNodeData.nodetype
                    else:
                        print 'childNode is not Ok.'

                # If this is a SearchResultsNode and there is at least one child, place the new entry above that first child.
                # This has the effect of putting Search Results nodes in Reverse Chronological Order.
                if (nodeType == 'SearchResultsNode') and childNode.IsOk():
                    insertPos = childNode

                # Complex comparison:
                #  1) is the childNode valid?
                #  2) does the child node's text match the next node in the nodelist?
                #  3) are the NodeTypes from compatible branches in the DB Tree?
                #  4) Is it a Keyword Example OR the correct record number

                # We're having a problem here going to Danish with UTF-8 translation files.  This is an attempt
                # to correct that problem.
                if ('unicode' in wx.PlatformInfo) and (type(node).__name__ == 'str'):
                    tmpNode = unicode(node, 'utf8')
                else:
                    tmpNode = node

                if DEBUG:
                    print "DatabaseTreeTab.add_Node(1)", childNode.IsOk()
#                    print "DatabaseTreeTab.add_Node(2)", itemText, node, itemText == node
                    if childNodeData != None:
                        print "DatabaseTreeTab.add_Node(3)", childNodeData.nodetype, expectedNodeType
                        print "DatabaseTreeTab.add_Node(4)", childNodeData.recNum, nodeRecNum

##                if childNode.IsOk():
##                    print childNode.IsOk(), (insertPos == None), (nodeListPos == len(nodeData) - 1), expectedNodeType, \
##                       itemText, node, (itemText > node), childNodeData.nodetype

                # Adding clips to the end of a long list was taking too long.  This code was added to make that process
                # go MUCH more quickly.  We look for the absence of an insertPos (so we're not needing to position a
                # clip in the middle of a list), reaching the end (n-1) of the nodeData, and where we're inserting a Clip.
                if (insertPos == None) and (nodeListPos == len(nodeData) - 1) and \
                   (expectedNodeType in ['QuoteNode', 'ClipNode', 'SnapshotNode']):
                    # Add the new Node to the Tree at the end
                    newNode = self.AppendItem(currentNode, node)
                    if expectedNodeType == 'QuoteNode':
                        # Add the tree node's graphic.
                        self.set_image(newNode, "Quote16")
                    elif expectedNodeType == 'ClipNode':
                        # Add the tree node's graphic.
                        self.set_image(newNode, "Clip16")
                    elif expectedNodeType == 'SnapshotNode':
                        # Add the tree node's graphic.
                        self.set_image(newNode, "Snapshot16")
                    # Get the current (clip) record number
                    currentRecNum = nodeRecNum
                    # Get the parent information
                    currentParent = nodeParent
                    # Use this data to create the node data
                    nodedata = _NodeData(nodetype=expectedNodeType, recNum=currentRecNum, parent=currentParent, sortOrder=sortOrder)
                    # Assign the node data to the new node.
                    self.SetPyData(newNode, nodedata)
                    # Signal that we're done!
                    notDone = False

                # if the text matches and the node type matches and we don't have a Keyword Example ...
                elif  (childNode.IsOk()) and \
                    (itemText.upper() == tmpNode.upper()) and \
                    (childNodeData.nodetype in expectedNodeType) and \
                    ((childNodeData.nodetype != 'KeywordExampleNode') or (childNodeData.recNum == nodeRecNum)):

                    if DEBUG:
                        print "In '%s', '%s' = '%s'.  Climbing on." % (self.GetItemText(currentNode), itemText, node)
                        print

                    # The indexPos gets incremented prematurely below, so set it to -1 so it increments properly to 0
                    indexPos = -1
                    
                    # We've found the next node.  Increment the nodeListPos counter.
                    nodeListPos += 1
                    expectedNodeType = self.UpdateExpectedNodeType(expectedNodeType, nodeListPos, nodeData, nodeType)

                    currentNode = childNode
                    notDone = False

                # Catch adding NEW Documents or Episodes and put them in the right spot alphabetically.
                elif (childNode.IsOk()) and \
                   (insertPos == None) and \
                   (((nodeListPos == len(nodeData) - 1) and \
                     (expectedNodeType in ['DocumentNode', 'EpisodeNode', 'SearchDocumentNode', 'SearchEpisodeNode']))   ) and \
                   ((itemText > node) or (childNodeData.nodetype == 'LibraryNoteNode')):
                    # Signal that we're done!
                    notDone = False
                    # Add the new Node to the Tree right BEFORE this Child Node
                    newNode = self.InsertItemBefore(currentNode, indexPos, node)
                    if expectedNodeType in ['DocumentNode', 'SearchDocumentNode']:
                        # Add the tree node's graphic.
                        self.set_image(newNode, "Document16")
                    elif expectedNodeType in ['EpisodeNode', 'SearchEpisodeNode']:
                        # Add the tree node's graphic.
                        self.set_image(newNode, "Episode16")
                        
                    # Get the current (clip) record number
                    currentRecNum = nodeRecNum
                    # Get the parent information
                    currentParent = nodeParent
                    # Use this data to create the node data
                    nodedata = _NodeData(nodetype=expectedNodeType, recNum=currentRecNum, parent=currentParent, sortOrder=sortOrder)
                    # Assign the node data to the new node.
                    self.SetPyData(newNode, nodedata)

                elif (not childNode.IsOk()) or (childNode == self.GetLastChild(currentNode)) or \
                     ((expectedNodeType == 'SearchResultsNode') and (nodeType == 'SearchResultsNode') and childNode.IsOk()):

                    if DEBUG:
                        print "Adding in '%s' to '%s'." % (node, self.GetItemText(currentNode))

                    if nodeListPos < len(nodeData) - 1:
                        
                        if DEBUG:
                            print "This is not the last Node.  It is node %s of %s" % (node, nodeData)

                        # Let's get the current Node's Data
                        currentNodeData = self.GetPyData(currentNode)

                        if DEBUG:
                            print "currentNode = %s, %s" % (self.GetItemText(currentNode), currentNodeData),

                        # We know what type of node is expected next, so we know what kind of node to add

                        # If the expected node is a Library...
                        if expectedNodeType == 'LibraryNode':
                            tempLibrary = Library.Library(node)
                            currentRecNum = tempLibrary.number

                        # If the expected node is a Episode...
                        elif expectedNodeType == 'EpisodeNode':
                            tempEpisode = Episode.Episode(series=Misc.unistrip(self.GetItemText(currentNode)), episode=node)
                            currentRecNum = tempEpisode.number

                        # If the expected node is a Transcript...
                        elif expectedNodeType == 'TranscriptNode':

                            print "ExpectedNodeType == 'TranscriptNode'.  This has not been coded!"

                        # If the expected node is a Document...
                        elif expectedNodeType == 'DocumentNode':

                            print "ExpectedNodeType == 'DocumentNode'.  This has not been coded!"

                        # If the expected node is a Collection...
                        elif expectedNodeType == 'CollectionNode':
                            tempCollection = Collection.Collection(node, currentNodeData.recNum)
                            currentRecNum = tempCollection.number

                        # If the expected node is a Clip...
                        elif expectedNodeType == 'ClipNode':

                            print "ExpectedNodeType == 'ClipNode'.  This has not been coded!"

                        # If the expected node is a Quote...
                        elif expectedNodeType == 'QuoteNode':

                            print "ExpectedNodeType == 'QuoteNode'.  This has not been coded!"

                        # If the expected node is a Snapshot...
                        elif expectedNodeType == 'SnapshotNode':

                            print "ExpectedNodeType == 'SnapshotNode'.  This has not been coded!"

                        # If the expected node is a Keyword Group...
                        elif expectedNodeType == 'KeywordGroupNode':
                            currentRecNum = 0

                        # If the expected node is a Keyword...
                        elif expectedNodeType == 'KeywordNode':
                            currentRecNum = 0

                        # if the expected node is a Keyword Example...
                        elif expectedNodeType == 'KeywordExampleNode':

                            print "ExpectedNodeType == 'KeywordExampleNode'.  This has not been coded!"

                        # If the expected node is a SearchLibrary...
                        elif expectedNodeType == 'SearchLibraryNode':
                            tempLibrary = Library.Library(node)
                            currentRecNum = tempLibrary.number

                        # If the LAST node is a SearchDocument...
                        elif expectedNodeType == 'SearchDocumentNode':

                            print "ExpectedNodeType == 'SearchDocumentNode'.  This has not been coded!"

                        # If the LAST node is a SearchEpisode...
                        elif expectedNodeType == 'SearchEpisodeNode':
                            tempEpisode = Episode.Episode(series=Misc.unistrip(self.GetItemText(currentNode)), episode=node)
                            currentRecNum = tempEpisode.number

                        # If the LAST node is a SearchTranscript...
                        elif expectedNodeType == 'SearchTranscriptNode':

                            print "ExpectedNodeType == 'SearchTranscriptNode'.  This has not been coded!"

                        # If the LAST node is a SearchCollection...
                        elif expectedNodeType == 'SearchCollectionNode':
                            tempCollection = Collection.Collection(node, currentNodeData.recNum)
                            currentRecNum = tempCollection.number

                        # If the LAST node is a SearchQuote...
                        elif expectedNodeType == 'SearchQuoteNode':

                            print "ExpectedNodeType == 'SearchQuoteNode'.  This has not been coded!"

                        # If the LAST node is a SearchClip...
                        elif expectedNodeType == 'SearchClipNode':

                            print "ExpectedNodeType == 'SearchClipNode'.  This has not been coded!"

                        # If the LAST node is a SearchSnapshot...
                        elif expectedNodeType == 'SearchSnapshotNode':

                            print "ExpectedNodeType == 'SearchSnapshotNode'.  This has not been coded!"

                        # The new node's parent record number will be the current node's record number
                        currentParent = currentNodeData.recNum

                    # If we ARE at the end of the Node List, we can use the values passed in by the calling routine
                    else:

                        if DEBUG:
                            print "end of the Node List"
                        
                        expectedNodeType = nodeType
                        currentRecNum = nodeRecNum
                        currentParent = nodeParent

                    if DEBUG:
                        print "Positioning in Tree...", insertPos, childNode.IsOk()
                    # We need to position new nodes in the proper place alphabetically and by nodetype.
                    # We can do this by setting the insertPos.  Note that we don't need to do this if
                    # we already have insertPos info based on Sort Order.
                    if (insertPos == None) and (childNode.IsOk()) and (childNode != self.GetLastChild(currentNode)):

                        # Let's get the current Node's Data
                        currentNodeData = self.GetPyData(currentNode)

                        if DEBUG:
                            print "nodetype = ", currentNodeData.nodetype, 'expectedNodeType = ', expectedNodeType
                            print

                        # Get the first Child Node of the Current Node
                        (child, cookieVal) = self.GetFirstChild(currentNode)
                        
                        if child.IsOk():
                            childData = self.GetPyData(child)

                            if nodeListPos < len(nodeData) - 1:
                                nt = expectedNodeType
                            else:
                                nt = nodeType
                            
                            if DEBUG:
                                print "DatabaseTreeTab.add_Node:",
                                print "Evaluate(%s, %s, %s, %s)" % (node, nt, self.GetItemText(child), childData)

                            while self.Evaluate(tmpNode, nt, child, childData):

                                (child, cookieVal) = self.GetNextChild(currentNode, cookieVal)
                                if child.IsOk():
                                    childData = self.GetPyData(child)
                                else:
                                    break

                            if child.IsOk():
                                insertPos = child
#                            else:
#                                if nodeType in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
#                                    sortOrder = childData.sortOrder + 1

                    # If no insertPos is specified, ...
                    if insertPos == None:
                        
                        # .. Add the new Node to the Tree at the end
                        newNode = self.AppendItem(currentNode, node)
                    # Otherwise ...
                    else:
                        # Get the first Child Node of the Current Node
                        (firstChild, cookieVal) = self.GetFirstChild(currentNode)
                        
                        if firstChild.IsOk():
                            # If our Insert Position is the First Child ...
                            if insertPos == firstChild:
                                # ... we need to "Prepend" the item to the parent's Nodes
                                newNode = self.PrependItem(currentNode, node)
                            # Otherwise ...
                            else:
                                # We want to insert BEFORE insertPos, not after it, so grab the Previous Sibling before doing the Insert!
                                insertPos = self.GetPrevSibling(insertPos)
                                # and then insert the item under the Parent, after the insertPos's Previous Sibling
                                newNode = self.InsertItem(currentNode, insertPos, node)
                        else:
                            newNode = self.AppendItem(currentNode, node)

                    # Give the new Node the appropriate Graphic
                    if (expectedNodeType == 'LibraryNode') or (expectedNodeType == 'SearchLibraryNode'):
                        self.set_image(newNode, "Library16")
                    elif (expectedNodeType == 'DocumentNode') or (expectedNodeType == 'SearchDocumentNode'):
                        self.set_image(newNode, "Document16")
                    elif (expectedNodeType == 'EpisodeNode') or (expectedNodeType == 'SearchEpisodeNode'):
                        self.set_image(newNode, "Episode16")
                    elif (expectedNodeType == 'TranscriptNode') or (expectedNodeType == 'SearchTranscriptNode'):
                        self.set_image(newNode, "Transcript16")
                    elif (expectedNodeType == 'CollectionNode') or (expectedNodeType == 'SearchCollectionNode'):
                        self.set_image(newNode, "Collection16")
                    elif (expectedNodeType == 'ClipNode') or (expectedNodeType == 'SearchClipNode') or (expectedNodeType == 'KeywordExampleNode'):
                        self.set_image(newNode, "Clip16")
                    elif (expectedNodeType == 'QuoteNode') or (expectedNodeType == 'SearchQuoteNode'):
                        self.set_image(newNode, "Quote16")
                    elif (expectedNodeType == 'SnapshotNode') or (expectedNodeType == 'SearchSnapshotNode'):
                        self.set_image(newNode, "Snapshot16")
                    elif expectedNodeType in ['NoteNode', 'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode',
                                              'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', \
                                              'SnapshotNoteNode', 'DocumentNoteNode', 'SearchNoteNode']:
                        self.set_image(newNode, "Note16")
                    elif expectedNodeType == 'KeywordGroupNode':
                        self.set_image(newNode, "KeywordGroup16")
                    elif expectedNodeType == 'KeywordNode':
                        self.set_image(newNode, "Keyword16")
                    elif expectedNodeType == 'SearchResultsNode':
                        self.set_image(newNode, "Search16")
                    else:
                      dlg = Dialogs.ErrorDialog(self, 'Undefined wxTreeCtrl.set_image() for nodeType %s in _DBTreeCtrl.add_Node()' % nodeType)
                      dlg.ShowModal()
                      dlg.Destroy()
                    # Create the Node Data and attach it to the Node
                    nodedata = _NodeData(nodetype=expectedNodeType, recNum=currentRecNum, parent=currentParent, sortOrder=sortOrder)
                    self.SetPyData(newNode, nodedata)

                    # Sort, if needed
                    if (insertPos == None) and \
                       (expectedNodeType in ['SearchLibraryNode']):
                        self.SortChildren(currentNode)
                    # If we're supposed to expand the node ...
                    if expandNode:
                        # ... expand it!
                        self.Expand(currentNode)

                    # We've found the next node.  Increment the nodeListPos counter.
                    nodeListPos += 1
                    # Get the Node Type we expect for the next node
                    expectedNodeType = self.UpdateExpectedNodeType(expectedNodeType, nodeListPos, nodeData, nodeType)
                    # update the Current Node with this new node's value
                    currentNode = newNode
                    # Signal that we're done
                    notDone = False
                    
                # Get the next child node
                (childNode, cookieItem) = self.GetNextChild(currentNode, cookieItem)
                indexPos += 1

        # If we've added a Clip, Quote, or Snapshot ...
        if nodeType in [ 'ClipNode', 'QuoteNode', 'SnapshotNode' ]:

#            print "DatabaseTreeTab.add_node():  Calling Sort Order", self.GetItemText(currentNode).encode('utf8')
#            print "sortOrder =", sortOrder, 'insertPos =', insertPos
            
            # When called through the ChatWindow, we run into a problem.  The Sort Orders in the tree's Item Data are
            # out of date.  We need to correct that here!
            if sortOrder != None:
                # Get the first item in the Collection's List
                (childNode, cookieItem) = self.GetFirstChild(currentNode)
                # Iterate through the Collection's Nodes as long as there are new items
                while childNode.IsOk():
                    # Get the child item's NodeData
                    tmpNodeData = self.GetPyData(childNode)

#                    print "  Before:", self.GetItemText(childNode).encode('utf8'), sortOrder, tmpNodeData.sortOrder, ' ==> ',
                    
                    # If the node data's Sort Order is equal to or greater than the Sort Order for the new item ...
                    if (tmpNodeData.sortOrder >= sortOrder):
                        # If the child node is a Clip ...
                        if tmpNodeData.nodetype == 'ClipNode':
                            # ... load the Clip's Data
                            tmpObj = Clip.Clip(tmpNodeData.recNum)
                            # Set the Node's Data to the Clip's correct Sort Order
                            tmpNodeData.sortOrder = tmpObj.sort_order
                            # and update the node's Node Data
                            self.SetPyData(childNode, tmpNodeData)
                        # If the child node is a Quote ...
                        elif tmpNodeData.nodetype == 'QuoteNode':
                            # ... load the Quote's Data
                            tmpObj = Quote.Quote(num=tmpNodeData.recNum)
                            # Set the Node's Data to the Quote's correct Sort Order
                            tmpNodeData.sortOrder = tmpObj.sort_order
                            # and update the node's Node Data
                            self.SetPyData(childNode, tmpNodeData)
                        # If the child node is a Snapshot ...
                        elif tmpNodeData.nodetype == 'SnapshotNode':
                            # ... load the Snapshot's Data
                            tmpObj = Snapshot.Snapshot(tmpNodeData.recNum)
                            # Set the Node's Data to the Snapshot's correct Sort Order
                            tmpNodeData.sortOrder = tmpObj.sort_order
                            # and update the node's Node Data
                            self.SetPyData(childNode, tmpNodeData)
                        # other data types, such as Nested Collections and Notes can be ignored

#                    print "  After:", sortOrder, tmpNodeData.sortOrder
                    
                    # Get the next child node
                    (childNode, cookieItem) = self.GetNextChild(currentNode, cookieItem)

        # If we've added a Clip, Quote, or Snapshot, or their Search equivalents ...
        if nodeType in [ 'LibraryNode', 'DocumentNode', 'EpisodeNode', 'ClipNode', 'QuoteNode', 'SnapshotNode',
                         'SearchLibraryNode', 'SearchDocumentNode', 'SearchEpisodeNode', 'SearchClipNode',
                         'SearchQuoteNode', 'SearchSnapshotNode' ]:
            # ... get the item's parent
            tmpNode = currentNode  # self.GetItemParent(currentNode)
            # ... and sort the parent
            self.SortChildren(tmpNode)
        # Refresh the Tree
        self.Refresh()

        # Calls from the MessagePost method of the Chat Window have caused exceptions.  This attempts to prevent that.
        if not avoidRecursiveYields:
            # There can be an issue with recursive calls to wxYield, so trap the exception ...
            try:
                wx.Yield()
            # ... and ignore it!
            except:
                pass

    def select_Node(self, nodeData, nodeType, ensureVisible=True):
        """ This method is used to select nodes in the tree.
            nodeData is a list that gives the tree structure that describes where the node should be selected. """

        currentNode = self.GetRootItem()

        # print "Root node = %s" % self.GetItemText(currentNode)
        # print "nodeData = ", nodeData

        # Having nodes and subnodes with the same name causes a variety of problems.  We need to track how far
        # down the tree branches we are to keep track of what object NodeTypes we should be working with.
        nodeListPos = 0
        if nodeType in ['LibraryRootNode', 'LibraryNode', 'EpisodeNode', 'TranscriptNode', 'DocumentNode', 
                        'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'DocumentNoteNode']:
            expectedNodeType = 'LibraryRootNode'
        elif nodeType in ['CollectionsRootNode', 'CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode',
                          'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']:
            expectedNodeType = 'CollectionsRootNode'
        elif nodeType in ['KeywordRootNode', 'KeywordGroupNode', 'KeywordNode', 'KeywordExampleNode']:
            expectedNodeType = 'KeywordRootNode'
        elif nodeType in ['SearchRootNode', 'SearchResultsNode', 'SearchLibraryNode', 'SearchDocumentNode', 'SearchEpisodeNode',
                          'SearchTranscriptNode', 'SearchCollectionNode', 'SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
            expectedNodeType = 'SearchRootNode'

        for node in nodeData:

            node = Misc.unistrip(node)

            # print "Looking for %s" % node

            notDone = True
            (childNode, cookieItem) = self.GetFirstChild(currentNode)
            while notDone and (currentNode != None):
                itemText = Misc.unistrip(self.GetItemText(childNode))
                # if childNode.IsOk():
                #     print "Looking in %s at %s for %s." % (self.GetItemText(currentNode), itemText, node)

                # Let's get the child Node's Data
                childNodeData = self.GetPyData(childNode)

                # To accept a node and climb on, the node text must match the text of the node being sought.
                # If the node being added is a Keyword Example Node, we also need to make sure that the node being
                # sought's Record Number matches the node being added's Record Number.  This is because it IS possible
                # to have two clips with the same name in different collections applied as examples of the same Keyword.

                # Complex comparison:
                #  1) is the childNode valid?
                #  2) does the child node's text match the next node in the nodelist?
                #  3) are the NodeTypes from compatible branches in the DB Tree?
                #    a) Library, Episode, or Transcript
                #    b) Collection or Clip 
                #    c) Keyword Group, Keyword, or Keyword Example
                #    d) Search Results Library, Episode, or Transcript
                #    e) Search Results Collection or Clip
                #  4) Is it a Keyword Example OR the correct record number

                # We're having a problem here going to Danish (and a couple other languages) with UTF-8 translation files.  
                # the following few lines were added to correct that problem.
                if ('unicode' in wx.PlatformInfo) and (type(node).__name__ == 'str'):
                    tmpNode = unicode(node, 'utf8')
                else:
                    tmpNode = node

                if DEBUG:
                    print "DatabaseTreeTab.add_Node(1)", childNode.IsOk()
                    # Can't print itemText or tmpNode, can't compare node!
                    print "DatabaseTreeTab.add_Node(2)", itemText.encode('utf8'), tmpNode.encode('utf8'), type(itemText), type(tmpNode), itemText == tmpNode
                    if childNodeData != None:
                        print "DatabaseTreeTab.add_Node(3)", childNodeData.nodetype, expectedNodeType

                if  (childNode.IsOk()) and \
                    \
                    (itemText.upper() == tmpNode.upper()) and \
                    \
                    (childNodeData.nodetype in expectedNodeType):

                    # We've found the next node.  Increment the nodeListPos counter.
                    nodeListPos += 1
                    expectedNodeType = self.UpdateExpectedNodeType(expectedNodeType, nodeListPos, nodeData, nodeType)

                    currentNode = childNode
                    notDone = False
                elif childNode == self.GetLastChild(currentNode):

                    if DEBUG:
                        dlg = Dialogs.ErrorDialog(self, 'Problem in _DBTreeCtrl.select_Node().\n"%s" not found for selection and display.' % node)
                        dlg.ShowModal()
                        dlg.Destroy()
                    currentNode = None
                    notDone = False

                if notDone:
                    (childNode, cookieItem) = self.GetNextChild(currentNode, cookieItem)
                
        if (currentNode != None) and ensureVisible:
            # Unselect whatever is currently selected
            self.UnselectAll()
            # Now select the item that is supposed to be selected ...
            self.SelectItem(currentNode)
            # ... and make sure that item is shown on screen
            self.EnsureVisible(currentNode)
        return currentNode

    def rename_Node(self, nodeData, nodeType, newName):
        """ This method is used to rename Nodes.  The single-user version of Transana was complete before
            the need for this method was discovered, so it may only be used by the Multi-user version. """
        # Select the appropriate Tree Node
        sel = self.select_Node(nodeData, nodeType, ensureVisible=False)
        # Change the Name of the Tree Node
        self.SetItemText(sel, newName)
        
        # Sort the Node
        # ... get the item's parent
        tmpNode = self.GetItemParent(sel)
        # ... and sort the parent
        self.SortChildren(tmpNode)

    def copy_Node(self, nodeType, sourceNodeData, destNodeData, deleteSourceNode, sendMessage=False):
        """ Copy or Move a tree node and all sub-nodes from the sourceNodeData location to the destNodeData location.
            NOTE:  This is NOT for Clip nodes, which require the Clip Sort Order concept. """
        # Get the actual source and destination NODES based on the node data passed in
        sourceNode = self.select_Node(sourceNodeData, nodeType, ensureVisible=False)
        destNode = self.select_Node(destNodeData, nodeType, ensureVisible=False)

        # Get the source node's Python Data
        pyData = self.GetPyData(sourceNode)
        # Get the Destination node's Python data
        destPyData = self.GetPyData(destNode)
        # Add the new Collection Node to the Tree
        self.add_Node(nodeType, destNodeData + (self.GetItemText(sourceNode), ), pyData.recNum, destPyData.recNum)
        # Get the actual node
        newNode = self.select_Node(destNodeData + (self.GetItemText(sourceNode),), nodeType, ensureVisible=False)
        # Keep a list of nodes to check, which may expand as we find nested nodes.  Start with the current
        # source node and the new node we just created as the destination.
        nodesToCheck = [(sourceNode, newNode)]
        # As long as there are nodes to check ...
        while len(nodesToCheck) > 0:
            # Get the new source and destination nodes
            (sourceNode, destNode) = nodesToCheck[0]
            # Remove those nodes from the list
            nodesToCheck = nodesToCheck[1:]
            # Get the new node's first child node
            (tmpNode, cookie) = self.GetFirstChild(sourceNode)
            # While there are child nodes to process...
            while tmpNode.IsOk():
                # ... add the new node to the databse tree with the name of the child node ...
                newNode = self.AppendItem(destNode, self.GetItemText(tmpNode))
                # ... get the child node's Python data ...
                pyData = self.GetPyData(tmpNode)
                # ... and get the destination node's Python data
                destPyData = self.GetPyData(destNode)
                # Assign the proper icon to the new node
                if pyData.nodetype in ['LibraryNode', 'SearchLibraryNode']:
                    self.set_image(newNode, 'Library16')
                elif pyData.nodetype in ['DocumentNode', 'SearchDocumentNode']:
                    self.set_image(newNode, 'Document16')
                elif pyData.nodetype in ['EpisodeNode', 'SearchEpisodeNode']:
                    self.set_image(newNode, 'Episode16')
                elif pyData.nodetype in ['TranscriptNode', 'SearchTranscriptNode']:
                    self.set_image(newNode, 'Transcript16')
                elif pyData.nodetype in ['CollectionNode', 'SearchCollectionNode']:
                    self.set_image(newNode, 'Collection16')
                elif pyData.nodetype in ['QuoteNode', 'SearchQuoteNode']:
                    self.set_image(newNode, 'Quote16')
                elif pyData.nodetype in ['ClipNode', 'SearchClipNode']:
                    self.set_image(newNode, 'Clip16')
                elif pyData.nodetype in ['SnapshotNode', 'SearchSnapshotNode']:
                    self.set_image(newNode, 'Snapshot16')
                elif pyData.nodetype in ['LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode',
                                         'ClipNoteNode', 'SnapshotNoteNode', 'DocumentNoteNode', 'QuoteNoteNode']:
                    self.set_image(newNode, 'Note16')
                # Set the new Parent in the Python data
                pyData.parent = destPyData.recNum
                # Set the Python data for the new node
                self.SetPyData(newNode, pyData)
                # If the child node has children of its own ...
                if self.ItemHasChildren(tmpNode):
                    # ... add it to the list of nodes to be processed
                    nodesToCheck.append((tmpNode, newNode))
                # Move on to the next child node
                (tmpNode, cookie) = self.GetNextChild(sourceNode, cookie)

        # If MU Messages should be sent (i.e. if this is called from anywhere BUT ChatWindow.py or
        # when copying or moving Search Results)
        if sendMessage:
              # Now let's communicate with other Transana instances if we're in Multi-user mode
              if not TransanaConstants.singleUserVersion:
                  # Prepare a Move Collection message
                  msg = "MCN %s"
                  # Convert the Node List to the form needed for messaging
                  data = ('Collections',)
                  for nd in sourceNodeData[1:]:
                       msg += " >|< %s"
                       data += (nd, )
                  # Send the message
                  if TransanaGlobal.chatWindow != None:
                      TransanaGlobal.chatWindow.SendMessage(msg % data)

        # If we're supposed to MOVE rather than COPY (i.e. delete the source node) ...
        if deleteSourceNode:
            # ... then delete the old Tree Node (This should not send the MU Message!)
            self.delete_Node(sourceNodeData, nodeType, sendMessage=False)

    def delete_Node(self, nodeData, nodeType, exampleClipNum=0, sendMessage=True, skipDelete=False):
        """ This method is used to delete nodes to the tree after it has been built.
            nodeData is a list that gives the tree structure that describes where the node should be deleted. """
        currentNode = self.GetRootItem()

        if DEBUG:
            print
            print "DatabaseTreeTab.delete_Node():"
            print "Root node = %s" % self.GetItemText(currentNode)
            print "nodeData = ", nodeData, "of type", nodeType
            print
        
        # See if we need to encode the first element of the tuple/list
        # (This shouldn't have to be so complicated!  But it's different on Mac-PPC and in different circumstances!)
        if type(_(nodeData[0])).__name__ == 'str':
            # We need to translate the root node name and convert it to Unicode.
            if type(nodeData) == type(()):
                nodeData = (unicode(_(nodeData[0]), 'utf8'),) + nodeData[1:]
            elif type(nodeData) == type([]):
                nodeData = [unicode(_(nodeData[0]), 'utf8')] + nodeData[1:]
        # Even if we don't need to encode it, we need to translate it
        else:
            # We need to translate the root node name and convert it to Unicode.
            if type(nodeData) == type(()):
                nodeData = (_(nodeData[0]),) + nodeData[1:]
            elif type(nodeData) == type([]):
                nodeData = [_(nodeData[0])] + nodeData[1:]

        # Having nodes and subnodes with the same name causes a variety of problems.  We need to track how far
        # down the tree branches we are to keep track of what object NodeTypes we should be working with.
        nodeListPos = 0
        if nodeType in ['LibraryRootNode', 'LibraryNode', 'EpisodeNode', 'TranscriptNode',
                        'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'DocumentNode', 'DocumentNoteNode']:
            expectedNodeType = 'LibraryRootNode'
        elif nodeType in ['CollectionsRootNode', 'CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode',
                          'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode']:
            expectedNodeType = 'CollectionsRootNode'
        elif nodeType in ['KeywordRootNode', 'KeywordGroupNode', 'KeywordNode', 'KeywordExampleNode']:
            expectedNodeType = 'KeywordRootNode'
        elif nodeType in ['SearchRootNode', 'SearchResultsNode', 'SearchLibraryNode', 'SearchDocumentNode', 'SearchEpisodeNode',
                          'SearchTranscriptNode', 'SearchCollectionNode', 'SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
            expectedNodeType = 'SearchRootNode'
        
        msgData = ''

        for node in nodeData:

            node = Misc.unistrip(node)

            if not TransanaConstants.singleUserVersion and sendMessage:
                if msgData != '':
                    msgData += ' >|< %s' % node
                else:
                    if 'unicode' in wx.PlatformInfo:
                        libraryPrompt = unicode(_('Libraries'), 'utf8')
                        collectionsPrompt = unicode(_('Collections'), 'utf8')
                        keywordsPrompt = unicode(_('Keywords'), 'utf8')
                    else:
                        # We need the nodeType as the first element.  Then, 
                        # we need the UNTRANSLATED label for the root node to avoid problems in mixed-language environments.
                        libraryPrompt = _('Libraries')
                        collectionsPrompt = _('Collections')
                        keywordsPrompt = _('Keywords')
                    if node == libraryPrompt:
                        msgData = nodeType + ' >|< Libraries'
                    elif node == collectionsPrompt:
                        msgData = nodeType + ' >|< Collections'
                    elif node == keywordsPrompt:
                        msgData = nodeType + ' >|< Keywords'

            notDone = True
            (childNode, cookieItem) = self.GetFirstChild(currentNode)

            while (childNode.IsOk()) and notDone and (currentNode != None):
                itemText = Misc.unistrip(self.GetItemText(childNode))
                childNodeData = self.GetPyData(childNode)

                # To accept a node and climb on, the node text must match the text of the node being sought.
                # If the node being added is a Keyword Example Node, we also need to make sure that the node being
                # sought's Record Number matches the node being added's Record Number.  This is because it IS possible
                # to have two clips with the same name in different collections applied as examples of the same Keyword.

                # Complex comparison:
                #  1) is the childNode valid?
                #  2) does the child node's text match the next node in the nodelist?
                #  3) are the NodeTypes from compatible branches in the DB Tree?
                #    a) Library, Episode, or Transcript
                #    b) Collection or Clip 
                #    c) Keyword Group, Keyword, or Keyword Example
                #    d) Search Results Library, Episode, or Transcript
                #    e) Search Results Collection or Clip
                #  4) Is it a Keyword Example OR the correct record number
                if  (childNode.IsOk()) and \
                    \
                    (itemText.upper() == node.upper()) and \
                    \
                    (childNodeData.nodetype in expectedNodeType) and \
                    \
                    ((exampleClipNum == 0) or (childNodeData.nodetype != 'KeywordExampleNode') or (childNodeData.recNum == exampleClipNum)):

                    # print "In %s, %s = %s.  Climbing on." % (self.GetItemText(currentNode), itemText, node)
                    
                    # We've found the next node.  Increment the nodeListPos counter.
                    nodeListPos += 1
                    expectedNodeType = self.UpdateExpectedNodeType(expectedNodeType, nodeListPos, nodeData, nodeType)

                    currentNode = childNode
                    notDone = False
                elif childNode == self.GetLastChild(currentNode):
                    if DEBUG:
                        dlg = Dialogs.ErrorDialog(self, 'Problem in _DBTreeCtrl.delete_Node().\n"%s" not found for delete.' % node)
                        dlg.ShowModal()
                        dlg.Destroy()
                    currentNode = None
                    notDone = False
                # if we're still looking, haven't yet reached the end of the list, ...
                if currentNode != None:
                    # ... get the next child node
                    (childNode, cookieItem) = self.GetNextChild(currentNode, cookieItem)

        # If we have found the correct node ...
        if currentNode != None:

            # See if (especially via MU Message) we need to clear the current interface because of the deletion

            # Get the data from the node to be deleted
            currentNodeData = self.GetPyData(currentNode)
            # If a Library is being deleted ...
            if nodeType == 'LibraryNode':
                # Get the Library Node's first Child
                (childNode, cookieItem) = self.GetFirstChild(currentNode)
                # As long as we have valid Child nodes ...
                while childNode.IsOk():
                    # ... get the Child Node's Data
                    childData = self.GetPyData(childNode)
                    # If the Child Node is an Episode ...
                    if childData.nodetype == 'EpisodeNode':
                        # ... we need to call this method recursively!
                        self.delete_Node(nodeData + (self.GetItemText(childNode), ), childData.nodetype, exampleClipNum, sendMessage, skipDelete=True)
                    elif childData.nodetype == 'DocumentNode':
                        # ... we need to remove it from the TranscriptWindow
                        self.parent.ControlObject.CloseOpenTranscriptWindowObject(Document.Document, childData.recNum)
                    # Get the next Child Node
                    (childNode, cookieItem) = self.GetNextChild(currentNode, cookieItem)

##                # If an Episode is currently loaded and
##                # the Library to be deleted contains the currently loaded Episode ...
##                if isinstance(self.parent.ControlObject.currentObj, Episode.Episode) and \
##                   (currentNodeData.recNum == self.parent.ControlObject.currentObj.series_num):
##                    # ... clear the interface!
##                    self.parent.ControlObject.ClearAllWindows()

            # If a Document is being deleted ...
            elif nodeType == 'DocumentNode':
                # If THIS Document is currently loaded, we need to remove it from the TranscriptWindow
                self.parent.ControlObject.CloseOpenTranscriptWindowObject(Document.Document, currentNodeData.recNum)

            # If an Episode is being deleted ...
            elif nodeType == 'EpisodeNode':
                # Get the Episode Node's first Child
                (childNode, cookieItem) = self.GetFirstChild(currentNode)
                # As long as we have valid Child nodes ...
                while childNode.IsOk():
                    # ... get the Child Node's Data
                    childData = self.GetPyData(childNode)
                    # If the Child Node is a Transcript ...
                    if childData.nodetype == 'TranscriptNode':
                        # ... we need to remove it from the TranscriptWindow
                        self.parent.ControlObject.CloseOpenTranscriptWindowObject(Transcript.Transcript, childData.recNum)
                    # Get the next Child Node
                    (childNode, cookieItem) = self.GetNextChild(currentNode, cookieItem)

            # If a Transcript is being deleted ...
            elif nodeType == 'TranscriptNode':
                # If THIS Transcript is currently loaded, we need to remove it from the TranscriptWindow
                self.parent.ControlObject.CloseOpenTranscriptWindowObject(Transcript.Transcript, currentNodeData.recNum)

            # If a Collection is being deleted ...
            elif nodeType == 'CollectionNode':
                # Get the Episode Node's first Child
                (childNode, cookieItem) = self.GetFirstChild(currentNode)
                # As long as we have valid Child nodes ...
                while childNode.IsOk():
                    # ... get the Child Node's Data
                    childData = self.GetPyData(childNode)
                    # If the Child Node is a Nested Collection ...
                    if childData.nodetype == 'CollectionNode':
                        # ... we need to call this method recursively!
                        self.delete_Node(nodeData + (self.GetItemText(childNode), ), childData.nodetype, exampleClipNum, sendMessage, skipDelete=True)
                    elif childData.nodetype == 'ClipNode':
                        # ... we need to remove it from the TranscriptWindow
                        self.parent.ControlObject.CloseOpenTranscriptWindowObject(Clip.Clip, childData.recNum)
                    elif childData.nodetype == 'QuoteNode':
                        # ... we need to remove it from the TranscriptWindow
                        self.parent.ControlObject.CloseOpenTranscriptWindowObject(Quote.Quote, childData.recNum)
                    elif childData.nodetype == 'SnapshotNode':
                        # iterate through all Snapshot Windows ...
                        for snapshotWindow in self.parent.ControlObject.SnapshotWindows:
                            # .. if THIS Snapshot is loaded ...
                            if snapshotWindow.obj.number == childData.recNum:
                                # ... close the snapshot window
                                snapshotWindow.Close()
                    # Get the next Child Node
                    (childNode, cookieItem) = self.GetNextChild(currentNode, cookieItem)

            # If a Quote is being deleted ...
            elif nodeType == 'QuoteNode':
                # If THIS Quote is currently loaded, we need to remove it from the TranscriptWindow
                self.parent.ControlObject.CloseOpenTranscriptWindowObject(Quote.Quote, currentNodeData.recNum)

            # If a Clip is being deleted ...
            elif nodeType == 'ClipNode':
                # If THIS Clip is currently loaded, we need to remove it from the TranscriptWindow
                self.parent.ControlObject.CloseOpenTranscriptWindowObject(Clip.Clip, currentNodeData.recNum)

            # If a Snapshot is being deleted ...
            elif nodeType == 'SnapshotNode':
                # iterate through all Snapshot Windows ...
                for snapshotWindow in self.parent.ControlObject.SnapshotWindows:
                    # .. if THIS Snapshot is loaded ...
                    if snapshotWindow.obj.number == currentNodeData.recNum:
                        # ... close the snapshot window
                        snapshotWindow.Close()

            # If a Keyword is being deleted ...
            elif nodeType == 'KeywordNode':

                kwg = self.GetItemText(self.GetItemParent(currentNode))
                kw = self.GetItemText(currentNode)

#                print
#                print "We have a keyword:  %s : %s" % (kwg, kw)
#                print "We should remove this keyword from ANY open object."
#                print

#                print "Iterate through open Notebook Pages"
                
                for page in range(self.parent.ControlObject.TranscriptWindow.nb.GetPageCount()):

                    tmpObj = self.parent.ControlObject.TranscriptWindow.nb.GetPage(page).GetChildren()[0].editor.TranscriptObj

#                    print self.parent.ControlObject.TranscriptWindow.nb.GetPageText(page), \
#                          type(tmpObj)
#                    print tmpObj
                    

                    if isinstance(tmpObj, Document.Document) or isinstance(tmpObj, Quote.Quote):

                        # .. if THIS object has THIS Keyword ...
                        if tmpObj.has_keyword(kwg, kw):
                            # ... remove the Keyword
                            tmpObj.remove_keyword(kwg, kw)

#                            print kwg, ':', kw, 'removed.'

#                    elif isinstance(tmpObj, Transcript.Transcript):

#                        print "We have an Episode or a Clip!  Maybe we don't have to do anything!"
#                    print
                        
#                print
                    

                # iterate through all Snapshot Windows ...
                for snapshotWindow in self.parent.ControlObject.SnapshotWindows:
                    # .. if THIS Snapshot has THIS Whole Snapshot Keyword ...
                    if snapshotWindow.obj.has_keyword(kwg, kw):
                        # ... remove the Keyword
                        snapshotWindow.obj.remove_keyword(kwg, kw)
                    # For each Detail Keyword in this Snapshot ...
                    for x in snapshotWindow.obj.codingObjects.keys():
                        # ... If the Detail Keyword is THIS Keyword ...
                        if (snapshotWindow.obj.codingObjects[x]['keywordGroup'] == kwg) and \
                           (snapshotWindow.obj.codingObjects[x]['keyword'] == kw):
                            # ... remove it from the Detail Coding
                            del(snapshotWindow.obj.codingObjects[x])
                    # If the Keyword Style is defined for this Snapshot ...
                    if (kwg, kw) in snapshotWindow.obj.keywordStyles.keys():
                        # ... remove the Keyword Style
                        del(snapshotWindow.obj.keywordStyles[(kwg, kw)])

                    # ... update the Snapshot Window
                    snapshotWindow.canvas.Freeze()
                    snapshotWindow.FileClear(None)
                    snapshotWindow.FileRedraw(None)
                    snapshotWindow.canvas.Thaw()

            if not skipDelete:
                # ... delete the current node
                self.Delete(currentNode)
            # If we are deleting a Keyword Example node here ...
            if (nodeType == 'KeywordExampleNode'):
                # ... add the example clip number to the end of the message, which gets used in ChatWindow.ProcessMessage to
                # ensure that the correct Keyword Example node gets removed on other copies of Transana.
                msgData = msgData + ' >|< %d' % exampleClipNum
            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion and sendMessage and (msgData != ''):
                if DEBUG:
                    print 'Message to send = "DN %s"' % msgData
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage("DN %s" % msgData)

    def UpdateCollectionSortOrder(self, node, sendMessage=True):
        """ Update the Sort Order Data for a Collection Node from the database and Sort the Node """
        # Get a dictionary of the Sort Order values from the database
        sortOrders = DBInterface.GetSortOrderData(self.GetPyData(node).recNum)
        # Iterate through the node's children
        # wxTreeCtrl requires the "cookie" value to list children.  Initialize it.
        cookie = 0
        # Get the first child of the dropNode 
        (tempNode, cookie) = self.GetFirstChild(node)
        # Iterate through all the node's children
        while tempNode.IsOk():
            # Get the current child's Node Data
            tempNodeData = self.GetPyData(tempNode)
            # If we have a Quote, Clip or a Snapshot ...
            if tempNodeData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
                # update the node's Sort Order
                tempNodeData.sortOrder = sortOrders[(tempNodeData.nodetype, tempNodeData.recNum)]
                # Save the node's updated data
                self.SetPyData(tempNode, tempNodeData)
            # If we are looking at the last Child in the Parent's Node, exit the while loop
            if tempNode == self.GetLastChild(node):
                break
            # If not, load the next Child record
            else:
                (tempNode, cookie) = self.GetNextChild(node, cookie)
        # Sort the Database Tree's Collection Node
        self.SortChildren(node)

        # If we're supposed to send the message and we're in MU,
        # let's send a message telling others they need to re-order this collection!
        if sendMessage and  not TransanaConstants.singleUserVersion:
            # Load the collection being re-ordered
            tempCollection = Collection.Collection(self.GetPyData(node).recNum)
            # Start by getting the Collection's Node Data
            nodeData = tempCollection.GetNodeData()
            # Indicate that we're sending a Collection Node
            msg = "CollectionNode"
            # Iterate through the nodes
            for node in nodeData:
                # ... add the appropriate seperator and then the node name
                msg += ' >|< %s' % node

            if DEBUG:
                print 'Message to send = "OC %s"' % msg

            # Send the Order Collection Node message
            if TransanaGlobal.chatWindow != None:
                TransanaGlobal.chatWindow.SendMessage("OC %s" % msg)

    def create_menus(self):
        """Create all the menu objects used in the tree control."""
        self.menu = {}          # Dictionary of menu references
        self.cmd_id_start = {}  # Dictionary of menu cmd_id starting points
        
        self.create_gen_menu()

        # Library Root Menu
        # Default Double-click is expand, then Add Library.  (See OnItemActivated())
        self.create_menu("LibraryRootNode",
                         (_("Add Library"),),
                         self.OnLibraryRootCommand)

        # Library Menu
        # Default Double-click is expand, then Add Episode.  (See OnItemActivated())
        tmpMenu = (_("Paste"),)
        if TransanaConstants.proVersion:
             tmpMenu += (_("Add Document"), )
        tmpMenu += (_("Add Episode"),)
        if TransanaConstants.proVersion:
            tmpMenu += (_("Batch Document Creation"),)
        tmpMenu += (_("Batch Episode Creation"), _("Add Library Note"), _("Delete Library"), _("Library Report"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Library Keyword Sequence Map"), _("Library Keyword Bar Graph"), _("Library Keyword Percentage Graph"))
        tmpMenu += (_("Analytic Data Export"), _("Library Properties"))
        self.create_menu("LibraryNode",
                         tmpMenu,
                         self.OnLibraryCommand)

        # Document Menu
        # Default Double-click is Open.  (See OnItemActivated())
        tmpMenu = (_("Cut"), _("Paste"), _("Open"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Open Additional Document"),)
        tmpMenu += (_("Add Document Note"), _("Delete Document"), _("Document Report"), _("Document Keyword Map"), _("Analytic Data Export"),
                    _("Document Properties"))
        self.create_menu("DocumentNode",
                         tmpMenu,
                         self.OnDocumentCommand)

        # Episode Menu
        # Default Double-click is expand, then Add Transcript.  (See OnItemActivated())
        tmpMenu = (_("Cut"), _("Paste"),
                   _("Add Transcript"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Open Multiple Transcripts"),)
        tmpMenu += (_("Add Episode Note"), _("Delete Episode"), _("Episode Report"), _("Episode Keyword Map"), _("Analytic Data Export"),
                    _("Episode Properties"))
        self.create_menu("EpisodeNode",
                         tmpMenu,
                         self.OnEpisodeCommand)

        # Transcript Menu
        # Default Double-click is Open.  (See OnItemActivated())
        tmpMenu = (_("Paste"), _("Open"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Open Additional Transcript"),)
        tmpMenu += (_("Add Transcript Note"), _("Delete Transcript"), _("Transcript Properties"))
        self.create_menu('TranscriptNode',
                         tmpMenu,
                         self.OnTranscriptCommand)

        # Collection Root Menu
        # Default Double-click is expand, then Add Collection.  (See OnItemActivated())
        self.create_menu('CollectionsRootNode',
                       (_("Paste"), _("Add Collection"), _("Collection Report"), _('Analytic Data Export')),
                        self.OnCollRootCommand)

        # Collection Menu
        # Default Double-click is expand, then Add Clip.  (See OnItemActivated())
        tmpMenu = (_("Cut"), _("Copy"), _("Paste"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Add Quote"),)
        tmpMenu += (_("Add Clip"),)
        if TransanaConstants.proVersion:
            tmpMenu += (_("Add Multi-transcript Clip"), _("Add Snapshot"), _("Batch Snapshot Creation"))
        tmpMenu += (_("Add Nested Collection"), _("Add Collection Note"), _("Delete Collection"),
                    _("Collection Report"), _("Collection Keyword Map"), _("Analytic Data Export"), _("Play All Clips"),
                    _("Collection Properties"))
        self.create_menu("CollectionNode",
                         tmpMenu,
                         self.OnCollectionCommand)

        # Quote Menu
        # Default Double-click is Open.  (See OnItemActivated())
        tmpMenu = (_("Cut"), _("Copy"), _("Paste"),
                   _("Open"), _("Add Quote"), _("Add Clip"), _("Add Multi-transcript Clip"), _("Add Snapshot"), _("Add Quote Note"),
                   _("Merge Quotes"), _("Delete Items"), _("Locate Quote in Document"), _("Insert Quote Hyperlink"), _("Quote Properties"),)
        self.create_menu('QuoteNode',
                         tmpMenu,
                         self.OnQuoteCommand)

        # Clip Menu
        # Default Double-click is Open.  (See OnItemActivated())
        tmpMenu = (_("Cut"), _("Copy"), _("Paste"),
                   _("Open"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Add Quote"),)
        tmpMenu += (_("Add Clip"),)
        if TransanaConstants.proVersion:
            tmpMenu += (_("Add Multi-transcript Clip"), _("Add Snapshot"),)
        tmpMenu += (_("Add Clip Note"), _("Merge Clips"),
                    _("Delete Items"), _("Locate Clip in Episode"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Export Clip Video"), _("Insert Clip Hyperlink"))
        tmpMenu += (_("Clip Properties"),)
        self.create_menu('ClipNode',
                         tmpMenu,
                         self.OnClipCommand)

        # Snapshot Menu
        # Default Double-click is Open.  (See OnItemActivated())
        self.create_menu('SnapshotNode',
                        (_("Cut"), _("Copy"), _("Paste"),
                         _("Open"), _("Add Quote"), _("Add Clip"), _("Add Multi-transcript Clip"), _("Add Snapshot"), 
                         _("Add Snapshot Note"), _("Delete Items"), _("Load Snapshot Context"), _("Insert Snapshot Hyperlink"),
                         _("Snapshot Properties")),
                         self.OnSnapshotCommand)

        # Keywords Root Menu
        # Default Double-click is expand, then Add Keyword Group.  (See OnItemActivated())
        self.create_menu("KeywordRootNode",
                        (_("Add Keyword Group"), _("Keyword Management"),
                         _("Keyword Summary Report")),
                        self.OnKwRootCommand)

        # Keyword Group Menu
        # Default Double-click is expand, then Add Keyword.  (See OnItemActivated())
        self.create_menu("KeywordGroupNode",
                        (_("Paste"),
                         _("Add Keyword"), _("Delete Keyword Group"), _("Keyword Summary Report")),
                        self.OnKwGroupCommand)

        # Keyword Menu
        # Default Double-click is expand, then Create Quick Clip.  (See OnItemActivated())
        tmpMenu = (_("Cut"), _("Copy"), _("Paste"),
                   _("Delete Keyword"))
        if TransanaConstants.proVersion:
            tmpMenu += (_("Create Quick Quote"),)
        tmpMenu += (_("Create Quick Clip"),)
        if TransanaConstants.proVersion:
            tmpMenu += (_("Create Multi-transcript Quick Clip"),)
        tmpMenu += (_("Quick Search"), _("Keyword Properties"))
        self.create_menu("KeywordNode",
                         tmpMenu,
                         self.OnKwCommand)

        # Keyword Example Menu
        # Default Double-click is Open.  (See OnItemActivated())
        self.create_menu("KeywordExampleNode",
                        (_("Open"), _("Locate this Clip"), _("Delete Keyword Example")),
                        self.OnKwExampleCommand)

        # Note Menu
        # Default Double-click is Open.  (See OnItemActivated())
        self.create_menu("NoteNode",
                        (_("Cut"), _("Copy"), _("Open"), _("Open in Notes Browser"), _("Delete Note"), _("Insert Note Hyperlink"),
                         _("Note Properties")),
                        self.OnNoteCommand)
        
        # The Search Root Node menu
        # Default Double-click is Search.  (See OnItemActivated())
        self.create_menu("SearchRootNode",
                         (_("Clear All"), _("Search")),
                         self.OnSearchCommand)
        
        # The Search Results Node Menu
        # Default Double-click is expand.  (See OnItemActivated())
        self.create_menu("SearchResultsNode",
                         (_("Paste"), _("Clear"), _("Convert to Collection"), _("Search Collection Report"), _("Play All Clips"),
                          _("Rename")),
                         self.OnSearchResultsCommand)
        
        # The Search Library Node Menu
        # Default Double-click is expand.  (See OnItemActivated())
        self.create_menu("SearchLibraryNode",
                        (_("Drop from Search Result"), _("Search Library Report")),
                        self.OnSearchLibraryCommand)
        
        # The Search Document Node Menu
        # Default Double-click is expand.  (See OnItemActivated())
        self.create_menu("SearchDocumentNode",
                        (_("Open"), _("Drop from Search Result"), _("Document Report"), _("Document Keyword Map")),
                        self.OnSearchDocumentCommand)
        
        # The Search Episode Node Menu
        # Default Double-click is expand.  (See OnItemActivated())
        self.create_menu("SearchEpisodeNode",
                        (_("Drop from Search Result"), _("Episode Report"), _("Episode Keyword Map")),
                        self.OnSearchEpisodeCommand)
        
        # The Search Transcript Node Menu
        # Default Double-click is Open.  (See OnItemActivated())
        self.create_menu("SearchTranscriptNode",
                         (_("Open"), _("Drop from Search Result")),
                         self.OnSearchTranscriptCommand)
        
        # The Search Collection Node Menu
        # Default Double-click is expand.  (See OnItemActivated())
        self.create_menu("SearchCollectionNode",
                        (_("Cut"), _("Copy"), _("Paste"),
                         _("Drop from Search Result"), _("Search Collection Report"),
                         _("Play All Clips"), _("Rename")),
                        self.OnSearchCollectionCommand)
        
        # The Search Quote Node Menu
        # Default Double-click is Open.  (See OnItemActivated())
        self.create_menu("SearchQuoteNode",
                        (_("Cut"), _("Copy"), _("Paste"),
                         _("Open"), _("Drop from Search Result"), _("Locate Quote in Document"),
                         _("Locate Quote in Collection"), _("Rename")),
                        self.OnSearchQuoteCommand)

        # The Search Clip Node Menu
        # Default Double-click is Open.  (See OnItemActivated())
        self.create_menu("SearchClipNode",
                        (_("Cut"), _("Copy"), _("Paste"),
                         _("Open"), _("Drop from Search Result"), _("Locate Clip in Episode"),
                         _("Locate Clip in Collection"), _("Rename")),
                        self.OnSearchClipCommand)

        # The Search Snapshot Node Menu
        # Default Double-click is Open.  (See OnItemActivated())
        self.create_menu('SearchSnapshotNode',
                        (_("Cut"), _("Copy"), _("Paste"),
                         _("Open"), _("Drop from Search Result"), _("Load Snapshot Context"), _("Locate Snapshot in Collection"),
                         _("Rename")),
                         self.OnSearchSnapshotCommand)


    def create_menu(self, name, items, handler):
        menu = wx.Menu()
        self.cmd_id_start[name] = self.cmd_id
        for item_s in items:
            menu.Append(self.cmd_id, item_s)
            wx.EVT_MENU(self, self.cmd_id, handler)
            self.cmd_id += 1
        self.menu[name] = menu

    def create_gen_menu(self):
        menu = wx.Menu()
        menu.Append(self.cmd_id, _("Update Database Window"))
        self.gen_menu = menu
        wx.EVT_MENU(self, self.cmd_id, self.refresh_tree)
        self.cmd_id += 1

    def OnLibraryRootCommand(self, evt):
        """Handle selections for root Library menu."""
        n = evt.GetId() - self.cmd_id_start["LibraryRootNode"]
        
        if n == 0:      # Add Library
            self.parent.add_series()
        else:
            raise MenuIDError
  
    def OnLibraryCommand(self, evt):
        """Handle menu selections for Library objects."""
        n = evt.GetId() - self.cmd_id_start["LibraryNode"]
        # If we're in the Standard version, we need to adjust the menu numbers
        # for Add Document (1), Batch Document Creation (3), Library Keyword Sequence Map (8), Library Keyword Bar Graph (9),
        # and Library Keyword Percentage Graph (10)
        if not TransanaConstants.proVersion:
            if (n >= 1):
                n += 1
            if (n >= 3):
                n += 1
            if (n >= 8):
                n += 3

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            library_name = self.GetItemText(sel)
        
        if n == 0:      # Paste
            # Do we need to propagate Keywords?  Assume NO!
            needToPropagate = False
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                
                # Multiple SOURCE items
                if isinstance(data, list):
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to all %s in %s "%s"?'), 'utf8')
                                data1 = unicode(_('Items'), 'utf8')
                                data2 = unicode(_('Libraries'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to all %s in %s "%s"?')
                                data1 = _('Items')
                                data2 = _('Libraries')
                            # Set up data to go with the prompt
                            promptdata = (data1, data2, self.GetItemText(sel))
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                            # Determine if we need to propagate keywords
                            needToPropagate = (result == wx.ID_YES)
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes (but we DON'T need to propagate Keywords!)
                            result = wx.ID_YES
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to all %s in multiple %s?'), 'utf8')
                                data1 = unicode(_('Items'), 'utf8')
                                data2 = unicode(_('Libraries'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to all %s in multiple %s?')
                                data1 = _('Items')
                                data2 = _('Libraries')
                            # Set up data to go with the prompt
                            promptdata = (data1, data2)
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                            # Determine if we need to propagate keywords
                            needToPropagate = (result == wx.ID_YES)
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes (but we DON'T need to propagate Keywords!)
                            result = wx.ID_YES
                # One SOURCE item
                else:
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # ... we skip a prompt and indicate the user said Yes
                        result = wx.ID_YES
                        # Prompt gets handled by DragAndDropObjects.DropKeyword().  We DO want user confirmation
                        confirmations = True
                        # If we have a Keyword Node ...
                        if data.nodetype == 'KeywordNode':
                            # ... we need to propagate keywords
                            needToPropagate = True
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a Keyword, get User confirmation
                        if data.nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add Keyword "%s:%s" to all %s in multiple %s?'), 'utf8') 
                                data1 = unicode(_('Items'), 'utf8')
                                data2 = unicode(_('Libraries'), 'utf8')
                            else:
                                prompt = _('Do you want to add Keyword "%s:%s" to all %s in multiple %s?')
                                data1 = _('Items')
                                data2 = _('Libraries')
                            # Set up data to go with the prompt
                            promptdata = (data.parent, data.text, data1, data2)
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(self.parent, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                            # Determine if we need to propagate keywords
                            needToPropagate = (result == wx.ID_YES)
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                        # We DO NOT want user confirmation
                        confirmations = False

                # If the user confirms, or we skip the confirmation programatically ...
                if result == wx.ID_YES:
                    # ... initialize a Keyword List
                    kwList = []
                    # If we need to propagate keywords, build the keyword list
                    if needToPropagate:
                        # If data is a list, there are multiple nodes to paste
                        if isinstance(data, list):
                            # For each keyword node item in the data list ...
                            for datum in data:
                                # ... create a Clip Keyword Object with the Keyword information ...
                                ckw = ClipKeywordObject.ClipKeyword(datum.parent, datum.text)
                                # ... and add the Clip Keyword Object to the Keyword List
                                kwList.append(ckw)
                        # If data is a single Keyword Node item ...
                        else:
                            # ... create a Clip Keyword Object with the Keyword information ...
                            ckw = ClipKeywordObject.ClipKeyword(data.parent, data.text)
                            # ... and add the Clip Keyword Object to the Keyword List
                            kwList.append(ckw)

                    # For each Library in the selected items ...
                    for item in selItems:
                        # ... grab the individual item ...
                        sel = item
                        # ... get the item's data ...
                        selData = self.GetPyData(sel)
                        # If data is a list, there are multiple Episode nodes to paste
                        if isinstance(data, list):
                            # If we need to propagate keywords ...
                            if needToPropagate:
                                # Now get a list of all Documents in the Library and iterate through them
                                for tempDocumentNum, tempDocumentID, tempLibraryNum in DBInterface.list_of_documents(selData.recNum):
                                    # ... propagating the new Document Keywords to all Quotes from that Document
                                    self.parent.ControlObject.PropagateObjectKeywords(_('Document'), tempDocumentNum, kwList)
                                # Now get a list of all Episodes in the Library and iterate through them
                                for tempEpisodeNum, tempEpisodeID, tempLibraryNum in DBInterface.list_of_episodes_for_series(self.GetItemText(sel)):
                                    # ... propagating the new Episode Keywords to all Clips from that Episode
                                    self.parent.ControlObject.PropagateObjectKeywords(_('Episode'), tempEpisodeNum, kwList)
                            # Iterate through the Episode nodes
                            for datum in data:
                                # ... and paste the data
                                DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'], confirmations=False)
                        # if data is NOT a list, it is a single Episode node to paste
                        else:
                            # If we're not displaying confirmation messages ...
                            if not confirmations:


                                # I don't think this EVER get's called!!  DKW 1 / 23 / 2015

                                
                                # If we need to propagate keywords ...
                                if needToPropagate:
                                    # Now get a list of all Documents in the Library and iterate through them
                                    for tempDocumentNum, tempDocumentID, tempLibraryNum in DBInterface.list_of_documents(selData.recNum):
                                        # ... propagating the new Document Keywords to all Quotes from that Document
                                        self.parent.ControlObject.PropagateObjectKeywords(_('Document'), tempDocumentNum, kwList)

##                                        print "Propagating One to ", tempDocumentID
                                        
                                    # Now get a list of all Episodes in the Library and iterate through them ...
                                    for tempEpisodeNum, tempEpisodeID, tempLibraryNum in DBInterface.list_of_episodes_for_series(self.GetItemText(sel)):
                                        # ... propagating the new Episode Keywords to all Clips from that Episode
                                        self.parent.ControlObject.PropagateObjectKeywords(_('Episode'), tempEpisodeNum, kwList)
                            # ... and paste the data
                            DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'], confirmations=confirmations)

                # Clear the Clipboard.  We can't paste again, since the data has been moved!
                DragAndDropObjects.ClearClipboard()
            # Close the Clipboard
            wx.TheClipboard.Close()


        elif n == 1:    # Add Document
            # Add a Document
            document_name = self.parent.add_document(library_name)
            # If the Document is successfully loaded ( != None) ...
            if document_name != None:
                # Select the Document in the Data Tree
                sel = self.select_Node((_('Libraries'), library_name, document_name), 'DocumentNode')
                # Load the newly created Document
                self.parent.ControlObject.LoadDocument(library_name, document_name, self.GetPyData(sel).recNum)  # Load everything via the ControlObject

        elif n == 2:    # Add Episode
            # Add an Episode
            episode_name = self.parent.add_episode(library_name)
            # If the Episode is successfully loaded ( != None) ...
            if episode_name != None:
                # Select the Episode in the Data Tree
                sel = self.select_Node((_('Libraries'), library_name, episode_name), 'EpisodeNode')
                # Automatically prompt to create a Transcript
                transcript_name = self.parent.add_transcript(library_name, episode_name)
                # If the Transcript is created ( != None) ...
                if transcript_name != None:
                    # Select the Transcript in the Data Tree
                    sel = self.select_Node((_('Libraries'), library_name, episode_name, transcript_name), 'TranscriptNode')
                    # Set the active transcript to 0 so the whole interface will be reset
                    self.parent.ControlObject.activeTranscript = 0
                    # Load the newly created Transcript
                    self.parent.ControlObject.LoadTranscript(library_name, episode_name, transcript_name)  # Load everything via the ControlObject

        elif n == 3:    # Batch Document Creation
            # Load the Library
            library = Library.Library(selData.recNum)
            try:
                # Lock the Library, to prevent it from being deleted out from under the Batch Document Creation
                library.lock_record()
                libraryLocked = True
            # Handle the exception if the record is already locked by someone else
            except RecordLockedError, s:
                # If we can't get a lock on the Library, it's really not that big a deal.  We only try to get it
                # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
                # still be able to add Documents even if someone else is editing the Library properties.
                libraryLocked = False
            # Create a dialog to get a list of media files to be processed
            fileListDlg = BatchFileProcessor.BatchFileProcessor(self, mode="document")
            # Show the dialog and get the list of files the user selects
            fileList = fileListDlg.get_input()
            # Close the dialog
            fileListDlg.Close()
            # Destroy the dialog
            fileListDlg.Destroy()

            # If the user pressed OK after selecting some files ...
            if fileList != None:
                # If there are more than 10 items in the list ...
                if len(fileList) > 10:
                    # ... Create a Progess Dialog ..
                    progress = wx.ProgressDialog(_("Batch Document Creation"), _("Creating Documents"), len(fileList), self)
                    # ... and a file counter to track progress
                    fileCount = 0
                # ... iterate through the file list
                for filename in fileList:
                    # If there are more than 10 items in the list ...
                    if len(fileList) > 10:
                        # ... update the progress dialog
                        progress.Update(fileCount)
                        # ... and increment the file counter
                        fileCount += 1
                    # Get the file's path
                    (path, filenameandext) = os.path.split(filename)
                    # Get the file's root name and extension
                    (filenameroot, extension) = os.path.splitext(filenameandext)
                    # We need to make sure only Text Files are processed!
                    if extension[1:] in TransanaConstants.documentFileTypes:
                        # Create a blank Document
                        tmpDocument = Document.Document()
                        # Name the Document after the root file name
                        tmpDocument.id = filenameroot
                        # Assign the Library number
                        tmpDocument.library_num = library.number
                        # Assign the imported document file name
                        tmpDocument.imported_file = filename
                        # Assign the import date
                        tmpDocument.import_date = datetime.datetime.now()
                        # If an XML file, an RTF file, or a TXT file was found ...
                        if os.path.exists(filename):
                            # ... start exception handling ...
                            try:
                                # Open the file
                                f = open(filename, "r")
                                # Read the file straight into the Document Text
                                tmpDocument.text = f.read()
                                # if we have a text file but no txt header ...
                                if (filename[-4:].lower() == '.txt') and (tmpDocument.text[:3].lower() != 'txt'):
                                    # ... add "txt" to the start of the file to signal that it's a text file
                                    tmpDocument.text = 'txt\n' + tmpDocument.text
                                # Close the file
                                f.close()
                            # If exceptions are raised ...
                            except:
                                # ... we don't need to do anything here.
                                # The consequence is probably that the Transcript Text will be blank.
                                pass
                        # Start exception handling
                        try:
                            # Save the new Document.  An exception will be generated if a Document or Episode with this name already exists.
                            tmpDocument.db_save()

                            # Build the Node List for the new Document
                            nodeData = (_('Libraries'), library.id, tmpDocument.id)
                            # Add the new Document to the database tree
                            self.add_Node('DocumentNode', nodeData, tmpDocument.number, library.number)
                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:

##                                print "DatabaseTreeTab.OnLibraryCommand():  Batch Document Creation.  Needs MU Message 'AD' confirmed"
                                
                                # Create an "Add Document" message
                                msg = "AD %s"
                                # Build the message details
                                data = (nodeData[1],)
                                for nd in nodeData[2:]:
                                    msg += " >|< %s"
                                    data += (nd, )
                                if DEBUG:
                                    print 'Message to send =', msg % data
                                # If there's a Chat window ...
                                if TransanaGlobal.chatWindow != None:
                                    # ... send the message
                                    TransanaGlobal.chatWindow.SendMessage(msg % data)

                        # If a Save Error was generated by the Snapshot ...
                        except SaveError, e:
                            # Build an error message
                            msg = _('Transana was unable to import file "%s"\nduring Batch Document Creation.\nA Document named "%s" already exists.')
                            # Make it Unicode
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Show the error message, then clean up.
                            dlg = Dialogs.ErrorDialog(self, msg % (filename, tmpDocument.id))
                            dlg.ShowModal()
                            dlg.Destroy()
                    else:
                        prompt = unicode('File "%s" could not be processed during Batch Document Creation.', 'utf8')
                        # Show the error message, then clean up.
                        dlg = Dialogs.ErrorDialog(self, prompt % (filename, ))
                        dlg.ShowModal()
                        dlg.Destroy()

                # If there are more than 10 items in the list ...
                if len(fileList) > 10:
                    # ... destroy the progress dialog
                    progress.Destroy()

            # Unlock the Library, if we locked it.
            if libraryLocked:
                library.unlock_record()

        elif n == 4:    # Batch Episode Creation
            # Load the Library
            library = Library.Library(selData.recNum)
            try:
                # Lock the Library, to prevent it from being deleted out from under the Batch Episode Creation
                library.lock_record()
                libraryLocked = True
            # Handle the exception if the record is already locked by someone else
            except RecordLockedError, s:
                # If we can't get a lock on the Library, it's really not that big a deal.  We only try to get it
                # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
                # still be able to add Episodes even if someone else is editing the Library properties.
                libraryLocked = False
            # Create a dialog to get a list of media files to be processed
            fileListDlg = BatchFileProcessor.BatchFileProcessor(self, mode="episode")
            # Show the dialog and get the list of files the user selects
            fileList = fileListDlg.get_input()
            # Close the dialog
            fileListDlg.Close()
            # Destroy the dialog
            fileListDlg.Destroy()
            # If the user pressed OK after selecting some files ...
            if fileList != None:
                # ... iterate through the file list
                for filename in fileList:
                    # Get the file's path
                    (path, filenameandext) = os.path.split(filename)
                    # Get the file's root name and extension
                    (filenameroot, extension) = os.path.splitext(filenameandext)
                    # We need to make sure only Media Files are processed!
                    if extension[1:] in TransanaConstants.mediaFileTypes:
                        # Create a blank Episode
                        tmpEpisode = Episode.Episode()
                        # Name the Episode after the root file name
                        tmpEpisode.id = filenameroot
                        # Assign the media file
                        tmpEpisode.media_filename = filename
                        # Assign the Library number
                        tmpEpisode.series_num = selData.recNum
                        # Start exception handling
                        try:
                            # Save the new Episode.  An exception will be generated if an Episode with this name already exists.
                            tmpEpisode.db_save()
                            # If no exception was generated and the save was successful, we need another level of exception
                            # handling to lock the new Episode
                            try:
                                # Lock the Episode, to prevent it from being deleted out from under the add.
                                tmpEpisode.lock_record()
                                episodeLocked = True
                            # Handle the exception if the record is already locked by someone else
                            except RecordLockedError, e:
                                # If we can't get a lock on the Episode, it's really not that big a deal.  We only try to get it
                                # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
                                # still be able to add Transcripts even if someone else is editing the Episode properties.
                                # NOTE:  This shouldn't be possible.  How could someone else lock an Episode the hasn't appeared yet?
                                episodeLocked = False
                            # Build the Node List for the new Episode
                            nodeData = (_('Libraries'), library.id, tmpEpisode.id)
                            # Add the new Episode to the database tree
                            self.add_Node('EpisodeNode', nodeData, tmpEpisode.number, library.number)
                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                if DEBUG:
                                    print 'Message to send = "AE %s >|< %s"' % (nodeData[-2], nodeData[-1])
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("AE %s >|< %s" % (nodeData[-2], nodeData[-1]))

                            # Now lets create an empty Transcript record
                            tmpTranscript = Transcript.Transcript()
                            # Name the Transcript after the media file root name
                            tmpTranscript.id = filenameroot
                            # Assign the new Episode's number 
                            tmpTranscript.episode_num = tmpEpisode.number
                            # Let's see if there is a Transcript file.  See if an XML file exists in the media file's
                            # directory with the same root file name
                            if os.path.exists(os.path.join(path, filenameroot + '.xml')):
                                # If so, let's remember that.
                                fname = os.path.join(path, filenameroot + '.xml')
                            # See if an RTF file exists in the media file's
                            # directory with the same root file name
                            elif os.path.exists(os.path.join(path, filenameroot + '.rtf')):
                                # If so, let's remember that.
                                fname = os.path.join(path, filenameroot + '.rtf')
                            # If there's no RTF file, see if a TXT file exists in the media file's
                            # directory with the same root file name
                            elif os.path.exists(os.path.join(path, filenameroot + '.txt')):
                                # If so, let's remember that.
                                fname = os.path.join(path, filenameroot + '.txt')
                            # If there's no TXT file, see if a default.xml file exists in the video root
                            # directory
                            elif os.path.exists(os.path.join(TransanaGlobal.configData.videoPath, 'default.xml')):
                                # If so, let's remember that.
                                fname = os.path.join(TransanaGlobal.configData.videoPath, 'default.xml')
                            # If there's no default.xml file, see if a default.rtf file exists in the video
                            # root directory
                            elif os.path.exists(os.path.join(TransanaGlobal.configData.videoPath, 'default.rtf')):
                                # If so, let's remember that.
                                fname = os.path.join(TransanaGlobal.configData.videoPath, 'default.rtf')
                            # If none of these files exists ...
                            else:
                                # ... then signal that no file was found by setting fname to None
                                fname = False
                            # If an XML file, an RTF file, or a TXT file was found ...
                            if fname:
                                # ... start exception handling ...
                                try:
                                    # Open the file
                                    f = open(fname, "r")
                                    # Read the file straight into the Transcript Text
                                    tmpTranscript.text = f.read()
                                    # if we have a text file but no txt header ...
                                    if (fname[-4:].lower() == '.txt') and (tmpTranscript.text[:3].lower() != 'txt'):
                                        # ... add "txt" to the start of the file to signal that it's a text file
                                        tmpTranscript.text = 'txt\n' + tmpTranscript.text
                                    # Close the file
                                    f.close()
                                # If exceptions are raised ...
                                except:
                                    # ... we don't need to do anything here.
                                    # The consequence is probably that the Transcript Text will be blank.
                                    pass
                            # Now save the new Transcript record.  Since we just created the Episode, we don't have to worry
                            # about a conflicting Transcript already existing.
                            tmpTranscript.db_save()
                            # Build the node data for the new Transcript record
                            nodeData = (_('Libraries'), library.id, tmpEpisode.id, tmpTranscript.id)
                            # Add the Transcript record to the Database tree
                            self.add_Node('TranscriptNode', nodeData, tmpTranscript.number, tmpEpisode.number)
                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                if DEBUG:
                                    print 'Message to send = "AT %s >|< %s >|< %s"' % (nodeData[-3], nodeData[-2], nodeData[-1])
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("AT %s >|< %s >|< %s" % (nodeData[-3], nodeData[-2], nodeData[-1]))
                        
                            # Unlock the Episode record
                            if episodeLocked:
                                tmpEpisode.unlock_record()
                        # If a Save Error was generated by the Episode ...
                        except SaveError, e:
                            # Build an error message
                            msg = _('Transana was unable to import file "%s"\nduring Batch Episode Creation.\nAn Episode named "%s" already exists.')
                            # Make it Unicode
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Show the error message, then clean up.
                            dlg = Dialogs.ErrorDialog(self, msg % (filename, tmpEpisode.id))
                            dlg.ShowModal()
                            dlg.Destroy()

                    else:
                        prompt = unicode('File "%s" could not be processed during Batch Episode Creation.', 'utf8')
                        # Show the error message, then clean up.
                        dlg = Dialogs.ErrorDialog(self, prompt % (filename, ))
                        dlg.ShowModal()
                        dlg.Destroy()


            # Unlock the Library, if we locked it.
            if libraryLocked:
                library.unlock_record()

        elif n == 5:    # Add Note
            self.parent.add_note(libraryNum=selData.recNum)
            
        elif n == 6:    # Delete
            # To delete a Library, the Notes Browser MUST be closed!
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Library Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # For each Library in the selected items ...
                for item in selItems:
                    # ... grab the individual item ...
                    sel = item
                    # ... and get the item's data
                    selData = self.GetPyData(sel)

                    # Load the Selected Library
                    library = Library.Library(selData.recNum)
                    # Get user confirmation of the Library Delete request
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Are you sure you want to delete Library "%s" and all related Documents, Episodes, Transcripts and Notes?'), 'utf8')
                    else:
                        prompt = _('Are you sure you want to delete Library "%s" and all related Documents, Episodes, Transcripts and Notes?')
                    dlg = Dialogs.QuestionDialog(self, prompt % (self.GetItemText(sel)))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    # If the user confirms the Delete Request...
                    if result == wx.ID_YES:
#                        # if current object is an Episode ...
#                        if isinstance(self.parent.ControlObject.currentObj, Episode.Episode):
                        # ... Start by clearing all current objects
                        self.parent.ControlObject.ClearAllWindows(clearAllTabs=True)
                        try:
                            # Try to delete the Library, initiating a Transaction
                            delResult = library.db_delete(1)
                            # If successful, remove the Library Node from the Database Tree
                            if delResult:
                                # Get the full Node Branch by climbing it to one level above the root
                                nodeList = (self.GetItemText(sel),)
                                while (self.GetItemParent(sel) != self.GetRootItem()):
                                    sel = self.GetItemParent(sel)
                                    nodeList = (self.GetItemText(sel),) + nodeList
                                    # print nodeList
                                # Call the DB Tree's delete_Node method.
                                self.delete_Node(nodeList, 'LibraryNode')
                                # ... and if the Notes Browser is open, ...
                                if self.parent.ControlObject.NotesBrowserWindow != None:
                                    # ... we need to CHECK to see if any notes were deleted.
                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
                        # Handle the RecordLocked exception, which arises when records are locked!
                        except RecordLockedError, e:
                            # Display the Exception Message
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('You cannot delete Library "%s".\n%s'), 'utf8')
                            else:
                                prompt = _('You cannot delete Library "%s".\n%s')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (library.id, e.explanation))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle the DeleteError exception, which is used to prevent orphaning of Clips.
                        except DeleteError, e:
                            # Display the Exception Message
                            errordlg = Dialogs.InfoDialog(None, e.reason)
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle other exceptions
                        except:
                            if DEBUG:
                                print "DatabaseTreeTab.OnLibraryCommand():  Delete Library"
                                import traceback
                                traceback.print_exc(file=sys.stdout)
                            # Display the Exception Message, allow "continue" flag to remain true
                            prompt = "%s : %s"
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(prompt, 'utf8')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                
        elif n == 7:    # Library Report
            # Call the Report Generator.  We pass the Library Name and want to show Keywords.
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Library Report"), 'utf8'),
                                            seriesName=library_name,
                                            showFile=True,
                                            showTime=True,
                                            showDocImportDate=True,
                                            showKeywords=True)

        elif n == 8:    # Library Map -- Sequence Mode
            LibraryMap.LibraryMap(self, unicode(_("Library Keyword Sequence Map"), 'utf8'), selData.recNum, library_name, 1, controlObject = self.parent.ControlObject)
            
        elif n == 9:    # Library Map -- Bar Graph Mode
            LibraryMap.LibraryMap(self, unicode(_("Library Keyword Bar Graph"), 'utf8'), selData.recNum, library_name, 2, controlObject = self.parent.ControlObject)
            
        elif n == 10:    # Library Map -- Percentage Mode
            LibraryMap.LibraryMap(self, unicode(_("Library Keyword Percentage Graph"), 'utf8'), selData.recNum, library_name, 3, controlObject = self.parent.ControlObject)

        elif n == 11:    # Analytic Data Export
            self.AnalyticDataExport(libraryNum = selData.recNum)
            
        elif n == 12:    # Library Properties
            library = Library.Library()
            # FIXME: Gracefully handle when we can't load the Library.
            # (yes, this can happen.  for example if another user changes
            # the name of it before the tree is refreshed.  then you'll get
            # a RecordNotFoundError.  Then we should auto-refresh.
            library.db_load_by_name(library_name)
            self.parent.edit_series(library)
            
        else:
            raise MenuIDError
 
    def OnDocumentCommand(self, evt):
        """Handle selections for Document menu."""
        n = evt.GetId() - self.cmd_id_start['DocumentNode']

        # If we're in the Standard version, we need to adjust the menu numbers
        # for Open Additional Document (3)
        if not TransanaConstants.proVersion and (n >= 3):
            n += 1

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            document_name = self.GetItemText(sel)

        if n == 0:    # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())

                # Multiple SOURCE items
                if isinstance(data, list):
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                                data1 = unicode(_('Document'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to %s "%s"?')
                                data1 = _('Document')
                            # Set up data to go with the prompt
                            promptdata = (data1, self.GetItemText(sel))
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to multiple %s?'), 'utf8')
                                data1 = unicode(_('Documents'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to multiple %s?')
                                data1 = _('Documents')
                            # Set up data to go with the prompt
                            promptdata = (data1,)
                            # Set up data to go with the prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                    # We DO NOT want user confirmation
                    confirmations = False
                # One SOURCE item
                else:
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # ... we skip a prompt and indicate the user said Yes
                        result = wx.ID_YES
                        # Prompt gets handled by DragAndDropObjects.DropKeyword().  We DO want user confirmation
                        confirmations = True
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a Keyword, get User confirmation
                        if data.nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add Keyword "%s:%s" to multiple %s?'), 'utf8') 
                                data1 = unicode(_('Documents'), 'utf8')
                            else:
                                prompt = _('Do you want to add Keyword "%s:%s" to multiple %s?')
                                data1 = _('Documents')
                            # Set up data to go with the prompt
                            promptdata = (data.parent, data.text, data1)
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(self.parent, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                        # We DO NOT want user confirmation
                        confirmations = False

                # If the user confirms, or we skip the confirmation programatically ...
                if result == wx.ID_YES:
                    # ... initialize a Keyword List
                    kwList = []
                    # If data is a list, there are multiple nodes to paste
                    if isinstance(data, list):
                        # For each item in the data list ...
                        for datum in data:
                            # ... create a Clip Keyword object with the keyword information ...
                            ckw = ClipKeywordObject.ClipKeyword(datum.parent, datum.text)
                            # ... and add the Clip Keyword Object to the Keyword List
                            kwList.append(ckw)
                    # If there's only ONE object in Data, we have a single Keyword Node object
                    else:
                        # ... create a Clip Keyword object with the keyword information ...
                        ckw = ClipKeywordObject.ClipKeyword(data.parent, data.text)
                        # ... and add the Clip Keyword Object to the Keyword List
                        kwList.append(ckw)
                    
                    # For each Document in the selected items ...
                    for item in selItems:
                        # ... grab the individual item ...
                        sel = item
                        # If data is a list, there are multiple nodes to paste
                        if isinstance(data, list):
                            # ... get tthe item's data ...
                            selData = self.GetPyData(sel)
                            # ... and propagate the Episode Keywords for this Episode.
                            self.parent.ControlObject.PropagateObjectKeywords(_('Document'), selData.recNum, kwList)
                            # Iterate through the nodes
                            for datum in data:
                                # ... and paste the data
                                DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'], confirmations=False)
                        # if data is NOT a list, it is a single node to paste
                        else:
                            # if we're not displaying confirmations ...
                            if not confirmations:


                                # I don't think this EVER gets called!!  DKW 1/23/2015

                                
                                # ... get the item's data ...
                                selData = self.GetPyData(sel)

##                                print "DatabaseTreeTab.OnDocumentCommand(2):  Paste:  Document Propagate Keywords to Quotes not implemented!"
                            
                                # ... and propagate the Document Keywords for this Document.  (Handled elsewhere if we are displaying
                                #     confirmations.)
                                self.parent.ControlObject.PropagateObjectKeywords(_('Document'), selData.recNum, kwList)
                            # ... and paste the data
                            DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'], confirmations=confirmations)
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 2:    # Open
            # This is functionally the same as double-clicking, although this way, we can open multiple Documents at once!
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Document

        elif n == 3:    # Open additional Document
            for sel in self.GetSelections():
                selData = self.GetPyData(sel)
                # Get the Library Name, which is the tree selection's parent's text
                library_name = self.GetItemText(self.GetItemParent(sel))
                # Open the document as an additional Document
                self.parent.ControlObject.OpenAdditionalDocument(selData.recNum, library_name)
            
        elif n == 4:    # Add Document Note
            # Add a Document Note
            self.parent.add_note(documentNum=selData.recNum)

        elif n == 5:    # Delete Document
            # If the Notes Browser is open ...
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of Close Notes Browser request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            # If we are to continue ...
            if result == wx.ID_YES:
                # For each Document in the selected items ...
                for item in selItems:
                    # ... grab the individual item ...
                    sel = item
                    # ... and get the item's data
                    selData = self.GetPyData(sel)

                    # Load the Selected Document
                    document = Document.Document(selData.recNum)
                    # Get user confirmation of the Document Delete request
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Are you sure you want to delete Document "%s" and all related Notes?'), 'utf8')
                    else:
                        prompt = _('Are you sure you want to delete Document "%s" and all related Notes?')
                    dlg = Dialogs.QuestionDialog(self, prompt % (self.GetItemText(sel)))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    # If the user confirms the Delete Request...
                    if result == wx.ID_YES:
                        try:
                            # Try to delete the Document, initiating a Transaction
                            delResult = document.db_delete(1)
                            # If successful, remove the Document Node from the Database Tree
                            if delResult:
                                # Get the full Node Branch by climbing it to one level above the root
                                nodeList = (self.GetItemText(sel),)
                                while (self.GetItemParent(sel) != self.GetRootItem()):
                                    sel = self.GetItemParent(sel)
                                    nodeList = (self.GetItemText(sel),) + nodeList
                                    # print nodeList
                                # Call the DB Tree's delete_Node method.
                                self.delete_Node(nodeList, 'DocumentNode')
                                # ... and if the Notes Browser is open, ...
                                if self.parent.ControlObject.NotesBrowserWindow != None:
                                    # ... we need to CHECK to see if any notes were deleted.
                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
                        # Handle the RecordLocked exception, which arises when records are locked!
                        except RecordLockedError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('You cannot delete Document "%s".\n%s'), 'utf8')
                            else:
                                prompt = _('You cannot delete Document "%s".\n%s')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (document.id, e.explanation))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle the DeleteError exception, which is used to prevent orphaning of Quotes.
                        except DeleteError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            errordlg = Dialogs.InfoDialog(None, e.reason)
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle other exceptions
                        except:
                            # Display the Exception Message, allow "continue" flag to remain true
                            prompt = "%s : %s"
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(prompt, 'utf8')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                            errordlg.ShowModal()
                            errordlg.Destroy()

        elif n == 6:    # Document Report
            # Call the Report Generator.  We pass the Library and Documetn names, and we want to show
            # Original File Names, Position data, and Keywords but not Comments, or Quote Notes by default.
            # (The point of including parameters set to false is that it triggers their availability as options.)
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Document Report"), 'utf8'),
                                            seriesName=self.GetItemText(self.GetItemParent(sel)),
                                            documentName=document_name,
                                            showHyperlink=True,
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showQuoteText=True,
##                                            showSnapshotImage=0,
##                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showQuoteNotes=False ) # ,
##                                            showSnapshotNotes=False)

        elif n == 7:    # Document Keyword Map
            
            self.DocumentKeywordMapReport(selData.recNum, self.GetItemText(self.GetItemParent(sel)), document_name)

        elif n == 8:    # Analytic Data Export
            self.AnalyticDataExport(documentNum = selData.recNum)
            
        elif n == 9:    # Document Properties
            # Load the Document Object
            document = Document.Document(selData.recNum)
            # Edit the Document Properties
            self.parent.edit_document(document)

        else:
            raise MenuIDError

    def OnEpisodeCommand(self, evt):
        """Handle menu selections for Episode objects."""
        n = evt.GetId() - self.cmd_id_start["EpisodeNode"]
        # If we're in the Standard version, we need to adjust the menu numbers
        # for Open Multiple Transcripts (3)
        if not TransanaConstants.proVersion and (n >= 3):
            n += 1
       
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            episode_name = self.GetItemText(sel)
            library_name = self.GetItemText(self.GetItemParent(sel))
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:      # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())

                # Multiple SOURCE items
                if isinstance(data, list):
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                                data1 = unicode(_('Episode'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to %s "%s"?')
                                data1 = _('Episode')
                            # Set up data to go with the prompt
                            promptdata = (data1, self.GetItemText(sel))
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to multiple %s?'), 'utf8')
                                data1 = unicode(_('Episodes'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to multiple %s?')
                                data1 = _('Episodes')
                            # Set up data to go with the prompt
                            promptdata = (data1,)
                            # Set up data to go with the prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                    # We DO NOT want user confirmation
                    confirmations = False
                # One SOURCE item
                else:
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # ... we skip a prompt and indicate the user said Yes
                        result = wx.ID_YES
                        # Prompt gets handled by DragAndDropObjects.DropKeyword().  We DO want user confirmation
                        confirmations = True
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a Keyword, get User confirmation
                        if data.nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add Keyword "%s:%s" to multiple %s?'), 'utf8') 
                                data1 = unicode(_('Episodes'), 'utf8')
                            else:
                                prompt = _('Do you want to add Keyword "%s:%s" to multiple %s?')
                                data1 = _('Episodes')
                            # Set up data to go with the prompt
                            promptdata = (data.parent, data.text, data1)
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(self.parent, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                        # We DO NOT want user confirmation
                        confirmations = False

                # If the user confirms, or we skip the confirmation programatically ...
                if result == wx.ID_YES:
                    # ... initialize a Keyword List
                    kwList = []
                    # If data is a list, there are multiple nodes to paste
                    if isinstance(data, list):
                        # For each item in the data list ...
                        for datum in data:
                            # ... create a Clip Keyword object with the keyword information ...
                            ckw = ClipKeywordObject.ClipKeyword(datum.parent, datum.text)
                            # ... and add the Clip Keyword Object to the Keyword List
                            kwList.append(ckw)
                    # If there's only ONE object in Data, we have a single Keyword Node object
                    else:
                        # ... create a Clip Keyword object with the keyword information ...
                        ckw = ClipKeywordObject.ClipKeyword(data.parent, data.text)
                        # ... and add the Clip Keyword Object to the Keyword List
                        kwList.append(ckw)
                    
                    # For each Episode in the selected items ...
                    for item in selItems:
                        # ... grab the individual item ...
                        sel = item
                        # If data is a list, there are multiple nodes to paste
                        if isinstance(data, list):
                            # ... get tthe item's data ...
                            selData = self.GetPyData(sel)
                            # ... and propagate the Episode Keywords for this Episode.
                            self.parent.ControlObject.PropagateObjectKeywords(_('Episode'), selData.recNum, kwList)
                            # Iterate through the nodes
                            for datum in data:
                                # ... and paste the data
                                DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'], confirmations=False)
                        # if data is NOT a list, it is a single node to paste
                        else:
                            # if we're not displaying confirmations ...
                            if not confirmations:
                                # ... get the item's data ...
                                selData = self.GetPyData(sel)
                                # ... and propagate the Episode Keywords for this Episode.  (Handled elsewhere if we are displaying
                                #     confirmations.)
                                self.parent.ControlObject.PropagateObjectKeywords(_('Episode'), selData.recNum, kwList)
                            # ... and paste the data
                            DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'], confirmations=confirmations)
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 2:    # Add Transcript
            # Add the Transcript
            transcript_name = self.parent.add_transcript(library_name, episode_name)
            # If the Transcript is created ( != None) ...
            if transcript_name != None:
                # Select the Transcript in the Data Tree
                sel = self.select_Node((_('Libraries'), library_name, episode_name, transcript_name), 'TranscriptNode')
                # Signal total interface refresh by setting the active trancript to 0
                self.parent.ControlObject.activeTranscript = 0
                # Load the newly created Transcript via the ControlObject
                self.parent.ControlObject.LoadTranscript(library_name, episode_name, transcript_name)

        elif n == 3:    # Open Multiple Transcripts
            # Initialize a Transcript window Counter
            trCount = 0
            # Get the first child of the Episode
            (trItem, cookie) = self.GetFirstChild(sel)
            # If we have at least one legit Transcript ...
            if trItem.IsOk():
                # Signal total interface refresh by setting the active trancript to 0
                self.parent.ControlObject.activeTranscript = 0
                # Load the first Transcript via the ControlObject
                self.parent.ControlObject.LoadTranscript(library_name, episode_name, self.GetItemText(trItem))
                # Increment the Transcript window Counter
                trCount += 1
            # while there are additional transcripts and we haven't exceeded the Max Transcript Count  and we have a Transcript node
            # and the LoadTranscript call didn't fails (signalled by the lack of a VideoFilename in the Control Object) ...
            while trItem.IsOk() and (trCount < TransanaConstants.maxTranscriptWindows) and \
                  (self.GetPyData(trItem).nodetype == 'TranscriptNode') and (self.parent.ControlObject.VideoFilename != ''):
                # ... get the next child of the Episode, the sibling of the Transcript.
                trItem = self.GetNextSibling(trItem)
                # If we have a legit item and it is a transcript ...
                if trItem.IsOk() and self.GetPyData(trItem).nodetype == 'TranscriptNode':
                    # ... open it as an additional transcript ...
                    self.parent.ControlObject.OpenAdditionalTranscript(self.GetPyData(trItem).recNum, library_name, episode_name)
                    # ... and increment the transcript window counter.
                    trCount += 1
            # When all the dust settles, let's set the focus to the first Transcript window.  1/2 second should do.
            wx.CallLater(500, self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().ActivatePanel, 0)

        elif n == 4:    # Add Note
            self.parent.add_note(episodeNum=selData.recNum)

        elif n == 5:    # Delete
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Episode Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # For each Episode in the selected items ...
                for item in selItems:
                    # ... grab the individual item ...
                    sel = item
                    # ... and get the item's data
                    selData = self.GetPyData(sel)

                    # Load the Selected Episode
                    episode = Episode.Episode(selData.recNum)
                    # Get user confirmation of the Episode Delete request
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Are you sure you want to delete Episode "%s" and all related Transcripts and Notes?\n(Please note that the video file associated with this Episode will not be deleted.)'), 'utf8')
                    else:
                        prompt = _('Are you sure you want to delete Episode "%s" and all related Transcripts and Notes?\n(Please note that the video file associated with this Episode will not be deleted.)')
                    dlg = Dialogs.QuestionDialog(self, prompt % (self.GetItemText(sel)))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    # If the user confirms the Delete Request...
                    if result == wx.ID_YES:
                        # Bring the Media File to the Front
                        self.parent.ControlObject.BringTranscriptToFront()
                        # if current object is THIS Episode ...
                        while (isinstance(self.parent.ControlObject.currentObj, Episode.Episode)) and (self.parent.ControlObject.currentObj.number == episode.number):
                            # Start by clearing all current objects
                            self.parent.ControlObject.ClearAllWindows()
                        try:
                            # Try to delete the Episode, initiating a Transaction
                            delResult = episode.db_delete(1)
                            # If successful, remove the Episode Node from the Database Tree
                            if delResult:
                                # Get the full Node Branch by climbing it to one level above the root
                                nodeList = (self.GetItemText(sel),)
                                while (self.GetItemParent(sel) != self.GetRootItem()):
                                    sel = self.GetItemParent(sel)
                                    nodeList = (self.GetItemText(sel),) + nodeList
                                    # print nodeList
                                # Call the DB Tree's delete_Node method.
                                self.delete_Node(nodeList, 'EpisodeNode')
                                # ... and if the Notes Browser is open, ...
                                if self.parent.ControlObject.NotesBrowserWindow != None:
                                    # ... we need to CHECK to see if any notes were deleted.
                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
                        # Handle the RecordLocked exception, which arises when records are locked!
                        except RecordLockedError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('You cannot delete Episode "%s".\n%s'), 'utf8')
                            else:
                                prompt = _('You cannot delete Episode "%s".\n%s')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (episode.id, e.explanation))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle the DeleteError exception, which is used to prevent orphaning of Clips.
                        except DeleteError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            errordlg = Dialogs.InfoDialog(None, e.reason)
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle other exceptions
                        except:
                            # Display the Exception Message, allow "continue" flag to remain true
                            prompt = "%s : %s"
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(prompt, 'utf8')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                
        elif n == 6:    # Episode Report Generator
            # Call the Report Generator.  We pass the Library and Episode names, and we want to show
            # File Names, Time data, Transcripts, and Keywords but not Comments, or Clip Notes by default.
            # (The point of including parameters set to false is that it triggers their availability as options.)
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Episode Report"), 'utf8'),
                                            seriesName=library_name,
                                            episodeName=episode_name,
                                            showHyperlink=True,
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showTranscripts=True,
                                            showSnapshotImage=0,
                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showClipNotes=False,
                                            showSnapshotNotes=False)

        elif n == 7:    # Keyword Map Report
            self.EpisodeKeywordMapReport(selData.recNum, library_name, episode_name)

        elif n == 8:    # Analytic Data Export
            self.AnalyticDataExport(episodeNum = selData.recNum)

        elif n == 9:    # Episode Properties
            library_name = self.GetItemText(self.GetItemParent(sel))
            episode = Episode.Episode()
            # FIXME: Gracefully handle when we can't load the Episode.
            episode.db_load_by_name(library_name, episode_name)
            # We can't edit properties for the currently-open Episode ...
            if isinstance(self.parent.ControlObject.currentObj, Episode.Episode) and (episode.number == self.parent.ControlObject.currentObj.number):
                # ... so clear all the windows if that is detected.
                self.parent.ControlObject.ClearAllWindows()
            # Call the Episode Properties screen
            self.parent.edit_episode(episode)

        else:
            raise MenuIDError

    def OnTranscriptCommand(self, evt):
        """ Handle menuy selections for Transcript menu """
        n = evt.GetId() - self.cmd_id_start['TranscriptNode']
        # If we're in the Standard version, we need to adjust the menu numbers
        # for Open Additional Transcript (2)
        if not TransanaConstants.proVersion and (n >= 2):
            n += 1
       
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            transcript_name = self.GetItemText(sel)
        
        if n == 0:      # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                # Multiple SOURCE items
                if isinstance(data, list):
                    # Iterate through the nodes
                    for datum in data:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                # One SOURCE item
                else:
                    # Process the Paste
                    DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
                # If we've MOVED rather than COPIED ...
                if self.cutCopyInfo['action'] == 'Move':
                    # ... Clear the Clipboard.  We can't paste again, since the data has been moved!
                    DragAndDropObjects.ClearClipboard()
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 1:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Transcript
            # If opening more than one Transcript ...
            if len(self.GetSelections()) > 1:
                # When all the dust settles, let's set the focus to the first Transcript window.  1/2 second should do.
                self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().ActivatePanel(0)
#                wx.CallLater(500, self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().ActivatePanel, 0)

        elif n == 2:    # Open additional Transcript
            for sel in self.GetSelections():
                selData = self.GetPyData(sel)
                # Get the Episode Name, which is the tree selection's parent's text
                episode_name = self.GetItemText(self.GetItemParent(sel))
                # Get the Library Name, which is the tree selection's grand-parent's text
                library_name = self.GetItemText(self.GetItemParent(self.GetItemParent(sel)))
                # Open the transcript as an additional Transcript
                self.parent.ControlObject.OpenAdditionalTranscript(selData.recNum, library_name, episode_name)
            
        elif n == 3:    # Add Note
            self.parent.add_note(transcriptNum=selData.recNum)

        elif n == 4:    # Delete
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Transcript Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # Bring the Transcript to the front
                self.parent.ControlObject.BringTranscriptToFront()
                # For each Transcript in the selected items ...
                for item in selItems:
                    # ... grab the individual item ...
                    sel = item
                    # ... and get the item's data
                    selData = self.GetPyData(sel)

                    # Load the Selected Transcript
                    # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                    transcript = Transcript.Transcript(selData.recNum, skipText=True)

                    # If the transcript is for the currently-loaded Episode:
                    while isinstance(self.parent.ControlObject.currentObj, Episode.Episode) and \
                       (self.parent.ControlObject.currentObj.number == selData.parent):
                        # Clear the interface before proceeding.  (in delete_node is too late!!)
                        self.parent.ControlObject.ClearAllWindows()

                    # Get user confirmation of the Transcript Delete request
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Are you sure you want to delete Transcript "%s" and all related Notes?'), 'utf8')
                    else:
                        prompt = _('Are you sure you want to delete Transcript "%s" and all related Notes?')
                    dlg = Dialogs.QuestionDialog(self, prompt % (self.GetItemText(sel)))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    # If the user confirms the Delete Request...
                    if result == wx.ID_YES:
                        try:
                            # Try to delete the Transcript, initiating a Transaction
                            delResult = transcript.db_delete(1)
                            # If successful, remove the Transcript Node from the Database Tree
                            if delResult:
                                # if current object is a Transcript in the current Episode ...
                                if (isinstance(self.parent.ControlObject.currentObj, Episode.Episode)) and \
                                   (selData.recNum in self.parent.ControlObject.TranscriptNum):
                                    # Start by clearing all current objects
                                    self.parent.ControlObject.ClearAllWindows()
                                # Get the full Node Branch by climbing it to one level above the root
                                nodeList = (self.GetItemText(sel),)
                                while (self.GetItemParent(sel) != self.GetRootItem()):
                                    sel = self.GetItemParent(sel)
                                    nodeList = (self.GetItemText(sel),) + nodeList
                                    # print nodeList
                                # Call the DB Tree's delete_Node method.
                                self.delete_Node(nodeList, 'TranscriptNode')
                                # ... and if the Notes Browser is open, ...
                                if self.parent.ControlObject.NotesBrowserWindow != None:
                                    # ... we need to CHECK to see if any notes were deleted.
                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
                        # Handle the RecordLocked exception, which arises when records are locked!
                        except RecordLockedError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('You cannot delete Transcript "%s".\n%s'), 'utf8')
                            else:
                                prompt = _('You cannot delete Transcript "%s".\n%s')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (transcript.id, e.explanation))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle the DeleteError exception, which is used to prevent orphaning of Clips.
                        except DeleteError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            errordlg = Dialogs.InfoDialog(None, e.reason)
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle other exceptions
                        except:
                            # Display the Exception Message, allow "continue" flag to remain true
                            prompt = "%s : %s"
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(prompt, 'utf8')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                            errordlg.ShowModal()
                            errordlg.Destroy()

        elif n == 5:    # Transcript Properties
            library_name = self.GetItemText(self.GetItemParent(self.GetItemParent(sel)))
            episode_name = self.GetItemText(self.GetItemParent(sel))
            episode = Episode.Episode()
            episode.db_load_by_name(library_name, episode_name)
            transcript = Transcript.Transcript(transcript_name, episode.number)
            self.parent.edit_transcript(transcript)

        else:
            raise MenuIDError

    def OnCollRootCommand(self, evt):
        """Handle selections for root Collection menu."""
        n = evt.GetId() - self.cmd_id_start['CollectionsRootNode']

        if n == 0:    # Paste
            # Get the list of selected items
            selItems = self.GetSelections()

            # specify the data formats to accept.
            #   Our data could be a DataTreeDragData object if the source is the Database Tree
            dfNode = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for these formats
            #   A DataTreeDragData object will populate the cdoNode object
            cdoNode = wx.CustomDataObject(dfNode)
            # Create a composite Data Object from the object types defined above
            cdo = wx.DataObjectComposite()
            cdo.Add(cdoNode)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get data from the Clipboard
            success = wx.TheClipboard.GetData(cdo)
            # If the data in the clipboard is in an appropriate format ...
            if success:
                # ... unPickle the data so it's in a usable form
                # First, let's try to get the DataTreeDragData object
                try:
                    data = cPickle.loads(cdoNode.GetData())
                except:
                    data = None
                    # If this fails, that's okay
                    pass

                # For each Collection in the selected items ...
                for item in selItems:
                    # ... grab the individual item ...
                    sel = item

                    # If data is a list, there are multiple Clip nodes to paste
                    if isinstance(data, list):
                        # Iterate through the Clip nodes
                        for datum in data:
                            # ... and paste the data
                            DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                    # if data is NOT a list, it is a single Clip node to paste
                    else:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 1:      # Add Collection
            self.parent.add_collection(0)

        elif n == 2:    # (Global) Collection Report
            # Call the Report Generator.  We pass an empty collection to signal the Global report. 
            # We want to show file names, clip time data, transcripts, and keywords but not Comments, collection
            # notes, or clip notes by default.  We also want to include 
            # nested collection data by default.  (The Global version is always empty without nested collections.)
            # (The point of including parameters set to false is that it triggers their availability as options.)
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Collection Report"), 'utf8'),
                                            collection=Collection.Collection(),
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showQuoteText=True,
                                            showTranscripts=True,
                                            showSnapshotImage=0,
                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showCollectionNotes=False,
                                            showQuoteNotes=False,
                                            showClipNotes=False,
                                            showSnapshotNotes=False,
                                            showNested=True,
                                            showHyperlink=True)

        elif n == 3:    # (Global) Analytic Data Export
            self.AnalyticDataExport()

        else:
            raise MenuIDError
    
    def OnCollectionCommand(self, evt):
        """Handle menu selections for Collection objects."""
        n = evt.GetId() - self.cmd_id_start['CollectionNode']
        # If we're in the Standard version, we need to adjust the menu numbers
        # for Add Quote (3), Add Multi-transcript Clip (5), Add Snapshot (6), and Batch Snapshot Creation (7)
        if not TransanaConstants.proVersion and (n >= 3):
            n += 1
        if not TransanaConstants.proVersion and (n >= 5):
            n += 3
        
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            coll_name = self.GetItemText(sel)
            parent_num = self.GetPyData(sel).parent
        
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            self.cutCopyInfo['sourceItem'] = sel
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            self.cutCopyInfo['sourceItem'] = sel
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:    # Paste
            # specify the data formats to accept.
            #   Our data could be a DataTreeDragData object if the source is the Database Tree
            dfNode = wx.CustomDataFormat('DataTreeDragData')

            # Specify the data object to accept data for these formats
            #   A DataTreeDragData object will populate the cdoNode object
            cdoNode = wx.CustomDataObject(dfNode)

            # Create a composite Data Object from the object types defined above
            cdo = wx.DataObjectComposite()
            cdo.Add(cdoNode)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get data from the Clipboard
            success = wx.TheClipboard.GetData(cdo)
            # If the data in the clipboard is in an appropriate format ...
            if success:
                # ... unPickle the data so it's in a usable form
                # First, let's try to get the DataTreeDragData object
                try:
                    data = cPickle.loads(cdoNode.GetData())
                except:
                    data = None
                    # If this fails, that's okay
                    pass

                # Multiple SOURCE items
                if isinstance(data, list):
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to all %s in %s "%s"?'), 'utf8')
                                data1 = unicode(_('Clips'), 'utf8')
                                data2 = unicode(_('Collection'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to all %s in %s "%s"?')
                                data1 = _('Clips and Snapshots')
                                data2 = _('Collection')
                            # Set up data to go with the prompt
                            promptdata = (data1, data2, self.GetItemText(sel))
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a list of Keywords, get User confirmation
                        if data[0].nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add multiple Keywords to all %s in multiple %s?'), 'utf8')
                                data1 = unicode(_('Clips and Snapshots'), 'utf8')
                                data2 = unicode(_('Collections'), 'utf8')
                            else:
                                prompt = _('Do you want to add multiple Keywords to all %s in multiple %s?')
                                data1 = _('Clips')
                                data2 = _('Collections')
                            # Set up data to go with the prompt
                            promptdata = (data1, data2)
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                # One SOURCE item
                else:
                    # One DESTINATION item
                    if len(selItems) == 1:
                        # ... we skip a prompt and indicate the user said Yes
                        result = wx.ID_YES
                        # Prompt gets handled by DragAndDropObjects.DropKeyword().  We DO want user confirmation
                        confirmations = True
                    # Multiple DESTINATION items
                    else:
                        # If the SOURCE is a Keyword, get User confirmation
                        if data.nodetype == 'KeywordNode':
                            # Get user confirmation of the Keyword Add request
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('Do you want to add Keyword "%s:%s" to all %s in multiple %s?'), 'utf8') 
                                data1 = unicode(_('Clips and Snapshots'), 'utf8')
                                data2 = unicode(_('Collections'), 'utf8')
                            else:
                                prompt = _('Do you want to add Keyword "%s:%s" to all %s in multiple %s?')
                                data1 = _('Clips')
                                data2 = _('Collections')
                            # Set up data to go with the prompt
                            promptdata = (data.parent, data.text, data1, data2)
                            # Show the user prompt
                            dlg = Dialogs.QuestionDialog(self.parent, prompt % promptdata)
                            result = dlg.LocalShowModal()
                            dlg.Destroy()
                        # If we don't have a Keyword ...
                        else:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                        # We DO NOT want user confirmation
                        confirmations = False

                # If the user confirms, or we skip the confirmation programatically ...
                if result == wx.ID_YES:                        
                    # For each Collection in the selected items ...
                    for item in selItems:
                        # ... grab the individual item ...
                        sel = item

                        # If data is a list, there are multiple Clip nodes to paste
                        if isinstance(data, list):
                            # Iterate through the Clip nodes
                            for datum in data:
                                # ... and paste the data
                                DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'], confirmations=False)
                        # if data is NOT a list, it is a single Clip node to paste
                        else:
                            # ... and paste the data
                            DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'], confirmations=confirmations)
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:    # Add Quote
            try:
                # Get the Document Selection information from the ControlObject.
                (documentNum, startChar, endChar, text) = self.parent.ControlObject.GetDocumentSelectionInfo()
                # If there's a selection in the text ...
                if text != '':
                    # ... copy it to the clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                    # Add the Quote.
                    self.parent.add_quote(coll_name)
                else:
                    # Create an error message.
                    msg = _('You must make a selection in a document to be able to add a Quote.')
                    errordlg = Dialogs.InfoDialog(None, msg)
                    errordlg.ShowModal()
                    errordlg.Destroy()
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

        elif n == 4:    # Add Clip
            try:
                # Get the Transcript Selection information from the ControlObject.
                (transcriptNum, startTime, endTime, text) = self.parent.ControlObject.GetTranscriptSelectionInfo()
                # If there's a selection in the text ...
                if text != '':
                    # ... copy it to the clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                # Add the Clip.
                self.parent.add_clip(coll_name)
            except:
                if DEBUG:
                    print sys.exc_info()[0]
                    print sys.exc_info()[1]
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                # Create an error message.
                msg = _('You must make a selection in a transcript to be able to add a Clip.')
                errordlg = Dialogs.InfoDialog(None, msg)
                errordlg.ShowModal()
                errordlg.Destroy()

        elif n == 5:    # Add Multi-transcript Clip
            # Create the appropriate Clip object
            tempClip = self.CreateMultiTranscriptClip(selData.recNum, coll_name)
            # Add the Clip, using EditClip because we don't want the overhead of the Drag-and-Drop architecture
            self.parent.edit_clip(tempClip)

        elif n == 6:    # Add Snapshot
            # Get the Collection (the Snapshot's Parent)
            coll = Collection.Collection(coll_name, parent_num)
            # Call the Add Snapshot dialog with the current Collection Number
            self.parent.add_snapshot(coll.number)

        elif n == 7:    # Batch Snapshot Creation
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # Load the Collection
            collection = Collection.Collection(selData.recNum)
            try:
                # Lock the Collection to prevent it from being deleted out from under the Batch Snapshot Creation
                collection.lock_record()
                collectionLocked = True
            # Handle the exception if the record is already locked by someone else
            except RecordLockedError, s:
                # If we can't get a lock on the Collection, it's really not that big a deal.  We only try to get it
                # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
                # still be able to add Snapshots even if someone else is editing the Collection properties.
                collectionLocked = False
            # Create a dialog to get a list of media files to be processed
            fileListDlg = BatchFileProcessor.BatchFileProcessor(self, mode="snapshot")
            # Show the dialog and get the list of files the user selects
            fileList = fileListDlg.get_input()
            # Close the dialog
            fileListDlg.Close()
            # Destroy the dialog
            fileListDlg.Destroy()
            # If the user pressed OK after selecting some files ...
            if fileList != None:
                # If there are more than 10 items in the list ...
                if len(fileList) > 10:
                    # ... Create a Progess Dialog ..
                    progress = wx.ProgressDialog(_("Batch Snapshot Creation"), _("Creating Snapshots"), len(fileList), self)
                    # ... and a file counter to track progress
                    fileCount = 0
                # ... iterate through the file list
                for filename in fileList:
                    # If there are more than 10 items in the list ...
                    if len(fileList) > 10:
                        # ... update the progress dialog
                        progress.Update(fileCount)
                        # ... and increment the file counter
                        fileCount += 1
                    # Get the file's path
                    (path, filenameandext) = os.path.split(filename)
                    # Get the file's root name and extension
                    (filenameroot, extension) = os.path.splitext(filenameandext)
                    # Create a blank Snapshot
                    tmpSnapshot = Snapshot.Snapshot()
                    # Name the Episode after the root file name
                    tmpSnapshot.id = filenameroot
                    # Assign the image file
                    tmpSnapshot.image_filename = filename
                    # Assign the Collection number and name
                    tmpSnapshot.collection_num = collection.number
                    tmpSnapshot.collection_id = collection.id
                    # Determine the Sort Order value
                    maxSortOrder = DBInterface.getMaxSortOrder(collection.number)
                    # Set the Sort Order
                    tmpSnapshot.sort_order = maxSortOrder + 1
                    # Start exception handling
                    try:
                        # Save the new Snapshot.  An exception will be generated if a Snapshot with this name already exists.
                        tmpSnapshot.db_save()

                        # Build the Node List for the new Snapshot
                        nodeData = (_('Collections'),) +  tmpSnapshot.GetNodeData(True)
                        # Add the new Snapshot to the database tree
                        self.add_Node('SnapshotNode', nodeData, tmpSnapshot.number, collection.number, sortOrder=tmpSnapshot.sort_order)
                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            # Create an "Add Snapshot" message
                            msg = "ASnap %s"
                            # Build the message details
                            data = (nodeData[1],)
                            for nd in nodeData[2:]:
                                msg += " >|< %s"
                                data += (nd, )
                            if DEBUG:
                                print 'Message to send =', msg % data
                            # If there's a Chat window ...
                            if TransanaGlobal.chatWindow != None:
                                # ... send the message
                                TransanaGlobal.chatWindow.SendMessage(msg % data)

                    # If a Save Error was generated by the Snapshot ...
                    except SaveError, e:
                        # Build an error message
                        msg = _('Transana was unable to import file "%s"\nduring Batch Snapshot Creation.\nA Snapshot named "%s" already exists.')
                        # Make it Unicode
                        if 'unicode' in wx.PlatformInfo:
                            msg = unicode(msg, 'utf8')
                        # Show the error message, then clean up.
                        dlg = Dialogs.ErrorDialog(self, msg % (filename, tmpSnapshot.id))
                        dlg.ShowModal()
                        dlg.Destroy()

                # If there are more than 10 items in the list ...
                if len(fileList) > 10:
                    # ... destroy the progress dialog
                    progress.Destroy()

            # Unlock the Collection, if we locked it.
            if collectionLocked:
                collection.unlock_record()

        elif n == 8:    # Add Nested Collection
            coll = Collection.Collection(coll_name, parent_num)
            self.parent.add_collection(coll.number)

        elif n == 9:    # Add Note
            self.parent.add_note(collectionNum=selData.recNum)

        elif n == 10:    # Delete
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Collection Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # For each Collection in the selected items ...
                for item in selItems:
                    # ... grab the individual item ...
                    sel = item
                    # ... and get the item's data
                    selData = self.GetPyData(sel)

                    # Load the Selected Collection
                    collection = Collection.Collection(selData.recNum)
                    # Get user confirmation of the Collection Delete request
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Are you sure you want to delete Collection "%s" and all related Nested Collections, Quotes, Clips and Notes?'), 'utf8')
                    else:
                        prompt = _('Are you sure you want to delete Collection "%s" and all related Nested Collections, Quotes, Clips and Notes?')
                    dlg = Dialogs.QuestionDialog(self, prompt % (self.GetItemText(sel)))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    # If the user confirms the Delete Request...
                    if result == wx.ID_YES:
                        # if current object is a Clip ...
                        if isinstance(self.parent.ControlObject.currentObj, Clip.Clip):
                            # Start by clearing all current objects
                            self.parent.ControlObject.ClearAllWindows()
                        # If the delete is successful, we will need to remove all Keyword Examples for all Clips that get deleted.
                        # Let's get a  list of those Keyword Examples before we do anything.
                        kwExamples = DBInterface.list_all_keyword_examples_for_all_clips_in_a_collection(selData.recNum)
                        try:
                            # Create a list starting with the CURRENT Collection
                            nestedCollectionList = [(selData.recNum, self.GetItemText(sel), selData.parent)]
                            # While the list has elements ...
                            while len(nestedCollectionList) > 0:
                                # ... add all nested Collections in the current Collection
                                (tmpColNum, tmpColId, tmpParentNum) = nestedCollectionList[0]
                                # ... drop the current Collection from the list 
                                nestedCollectionList = nestedCollectionList[1:] + DBInterface.list_of_collections(tmpColNum)
                                # ... get all the quotes in the current Collection
                                quoteList = DBInterface.list_of_quotes_by_collectionnum(tmpColNum)
                                # ... for each Quote ...
                                for quote in quoteList:
                                    # Remove the Object's Position Data from the Source Document, if it's open
                                    self.parent.ControlObject.RemoveQuoteFromOpenDocument(quote[0], quote[3])
                                    # Even if this computer doesn't need to update the Source Document, others might need to.
                                    if not TransanaConstants.singleUserVersion:
                                        # We need to pass the type of the deleted Object's record number and the deleted Object's
                                        # SOURCE Document number.
                                        
                                        if DEBUG:
                                            print 'Message to send = "DQPOD %s %s"' % (quote[0], quote[3])
                                            
                                        if (TransanaGlobal.chatWindow != None) and (quote[3] > 0):
                                            TransanaGlobal.chatWindow.SendMessage("DQPOD %s %s" % (quote[0], quote[3]))

                            # Try to delete the Collection, initiating a Transaction
                            delResult = collection.db_delete(1)
                            # If successful, remove the Collection Node from the Database Tree
                            if delResult:
                                # We need to remove the Keyword Examples from the Database Tree before we remove the Collection Node.
                                # Deleting all these ClipKeyword records needs to remove Keyword Example Nodes in the DBTree.
                                # That needs to be done here in the User Interface rather than in the Clip Object, as that is
                                # a user interface issue.  The Clip Record and the Clip Keywords Records get deleted, but
                                # the user interface does not get cleaned up by deleting the Clip Object.
                                for (kwg, kw, clipNum, clipID) in kwExamples:
                                    self.delete_Node((_("Keywords"), kwg, kw, clipID), 'KeywordExampleNode', exampleClipNum = clipNum)
                                
                                # Get the full Node Branch by climbing it to one level above the root
                                nodeList = (self.GetItemText(sel),)
                                while (self.GetItemParent(sel) != self.GetRootItem()):
                                    sel = self.GetItemParent(sel)
                                    nodeList = (self.GetItemText(sel),) + nodeList
                                    # print nodeList
                                # Call the DB Tree's delete_Node method.
                                self.delete_Node(nodeList, 'CollectionNode')
                                # ... and if the Notes Browser is open, ...
                                if self.parent.ControlObject.NotesBrowserWindow != None:
                                    # ... we need to CHECK to see if any notes were deleted.
                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
                                # Check to see if we need to update the keyword Visualization.  (We don't know if any clips
                                # were in the current episode, so all we can check is that an episode is loaded!)
                                if (isinstance(self.parent.ControlObject.currentObj, Episode.Episode) or
                                    isinstance(self.parent.ControlObject.currentObj, Document.Document)):
                                    self.parent.ControlObject.UpdateKeywordVisualization()
                                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                                if not TransanaConstants.singleUserVersion:
                                    # When a collection gets deleted, clips from anywhere could go with it.  It's safest
                                    # to update the Keyword Visualization no matter what.
                                    if DEBUG:
                                        print 'Message to send = "UKV %s %s %s"' % ('None', 0, 0)
                                        
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))
                                
                        # Handle the RecordLocked exception, which arises when records are locked!
                        except RecordLockedError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('You cannot delete Collection "%s".\n%s'), 'utf8')
                            else:
                                prompt = _('You cannot delete Collection "%s".\n%s')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (collection.id, e.explanation))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle other exceptions
                        except:
                            # Display the Exception Message, allow "continue" flag to remain true
                            prompt = "%s : %s"
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(prompt, 'utf8')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                            errordlg.ShowModal()
                            errordlg.Destroy()

        elif n == 11:    # Collection Report

            # Get the appropriate collection
            coll = Collection.Collection(coll_name, parent_num)
            # Call the Report Generator.  We pass the Collection Object, and we want to show
            # file names, clip time data, transcripts, and keywords but not Comments, collection notes, or
            # clip notes by default.  We also want to include nested collection data by default.
            # (The point of including parameters set to false is that it triggers their availability as options.)
            report = ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Collection Report"), 'utf8'),
                                            collection=coll,
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showQuoteText=True,
                                            showTranscripts=True,
                                            showSnapshotImage=0,
                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showCollectionNotes=False,
                                            showQuoteNotes=False,
                                            showClipNotes=False,
                                            showSnapshotNotes=False,
                                            showNested=True,
                                            showHyperlink=True)

        elif n == 12:    # Collection Keyword Map Report
            # Call the Collection Keyword Map 
            self.CollectionKeywordMapReport(selData.recNum)

        elif n == 13:    # Analytic Data Export
            # Call Analytic Data Export with the Collection Number
            self.AnalyticDataExport(collectionNum = selData.recNum)

        elif n == 14:    # Play All Clips
            # Get the appropriate collection
            coll = Collection.Collection(coll_name, parent_num)
            # Play All Clips takes the current Collection and the ControlObject as parameters.
            # (The ControlObject is owned not by the _DBTreeCtrl but by its parent)
            PlayAllClips.PlayAllClips(collection=coll, controlObject=self.parent.ControlObject)
            # Return the Data Window to the Database Tab
            self.parent.ControlObject.ShowDataTab(0)
            # Let's Update the Play State, so that if we've been in Presentation Mode, the screen will be reset.
            self.parent.ControlObject.UpdatePlayState(TransanaConstants.MEDIA_PLAYSTATE_STOP)
            # Let's clear all the Windows, since we don't want to stay in the last Clip played.
            self.parent.ControlObject.ClearAllWindows()

        elif n == 15:    # Collection Properties
            # FIXME: Gracefully handle when we can't load the Collection.
            coll = Collection.Collection(coll_name, parent_num)
            self.parent.edit_collection(coll)

        else:
            raise MenuIDError

    def OnQuoteCommand(self, evt):
        """Handle selections for Quote menu."""
        n = evt.GetId() - self.cmd_id_start['QuoteNode']

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            document_name = self.GetItemText(sel)
            coll_name = self.GetItemText(self.GetItemParent(sel))
            coll_parent_num = self.GetPyData(self.GetItemParent(sel)).parent

        if n == 0:    # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:  # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:  # Paste
            # specify the data formats to accept.
            #   Our data could be a DataTreeDragData object if the source is the Database Tree
            dfNode = wx.CustomDataFormat('DataTreeDragData')
            #   Our data could be a QuoteDragDropData object if the source is the Document (quote creation)
            dfQuote = wx.CustomDataFormat('QuoteDragDropData')
            #   Our data could be a ClipDragDropData object if the source is the Transcript (clip creation)
            dfClip = wx.CustomDataFormat('ClipDragDropData')

            # Specify the data object to accept data for these formats
            #   A DataTreeDragData object will populate the cdoNode object
            cdoNode = wx.CustomDataObject(dfNode)
            #   A QuoteDragDropData object will populate the cdoQuote object
            cdoQuote = wx.CustomDataObject(dfQuote)
            #   A ClipDragDropData object will populate the cdoClip object
            cdoClip = wx.CustomDataObject(dfClip)

            # Create a composite Data Object from the object types defined above
            cdo = wx.DataObjectComposite()
            cdo.Add(cdoNode)
            cdo.Add(cdoQuote)
            cdo.Add(cdoClip)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get data from the Clipboard
            success = wx.TheClipboard.GetData(cdo)
            # If the data in the clipboard is in an appropriate format ...
            if success:
                # ... unPickle the data so it's in a usable form
                # First, let's try to get the DataTreeDragData object
                try:
                    data = cPickle.loads(cdoNode.GetData())
                except:
                    data = None
                    # If this fails, that's okay
                    pass

                # Let's also try to get the QuoteDragDropData object
                try:
                    data2 = cPickle.loads(cdoQuote.GetData())
                except:
                    data2 = None
                    # If this fails, that's okay
                    pass

                # Let's also try to get the ClipDragDropData object
                try:
                    data3 = cPickle.loads(cdoClip.GetData())
                except:
                    data3 = None
                    # If this fails, that's okay
                    pass

                # if we didn't get the DataTreeDragData object, we need to substitute the QuoteDragDropData object
                if data == None:
                    data = data2
                # if we didn't get the DataTreeDragData or QuoteDragDropData object, we need to substitute the ClipDragDropData object
                if data == None:
                    data = data3
                    
                # If our clipboard data is text for creating a Quote ...
                if type(data) == type(DragAndDropObjects.QuoteDragDropData()):
                    # ... then create a Quote!  (I don't think this ever gets called!??)
                    DragAndDropObjects.CreateQuote(data, selData, self, sel)
                # If our clipboard data is text for creating a clip ...
                elif type(data) == type(DragAndDropObjects.ClipDragDropData()):
                    # ... then create a clip!  (I don't think this ever gets called!??)
                    DragAndDropObjects.CreateClip(data, selData, self, sel)
                # If we have drag data fron the Database Tree ...
                else:
                    # Multiple SOURCE items
                    if isinstance(data, list):
                        # One DESTINATION item
                        if len(selItems) == 1:
                            # If the SOURCE is a list of Keywords, get User confirmation
                            if data[0].nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                                    data1 = unicode(_('Quote'), 'utf8')
                                else:
                                    prompt = _('Do you want to add multiple Keywords to %s "%s"?')
                                    data1 = _('Quote')
                                # Set up data to go with the prompt
                                promptdata = (data1, self.GetItemText(sel))
                                # Show the user prompt
                                dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                        # Multiple DESTINATION items
                        else:
                            # If the SOURCE is a list of Keywords, get User confirmation
                            if data[0].nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add multiple Keywords to multiple %s?'), 'utf8')
                                    data1 = unicode(_('Quotes'), 'utf8')
                                else:
                                    prompt = _('Do you want to add multiple Keywords to multiple %s?')
                                    data1 = _('Quotes')
                                # Set up data to go with the prompt
                                promptdata = (data1,)
                                # Set up data to go with the prompt
                                dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                    # One SOURCE item
                    else:
                        # One DESTINATION item
                        if len(selItems) == 1:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                            # Prompt gets handled by DragAndDropObjects.DropKeyword().  We DO want user confirmation
                            confirmations = True
                        # Multiple DESTINATION items
                        else:
                            # If the SOURCE is a Keyword, get User confirmation
                            if data.nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add Keyword "%s:%s" to multiple %s?'), 'utf8') 
                                    data1 = unicode(_('Quotes'), 'utf8')
                                else:
                                    prompt = _('Do you want to add Keyword "%s:%s" to multiple %s?')
                                    data1 = _('Quotes')
                                # Set up data to go with the prompt
                                promptdata = (data.parent, data.text, data1)
                                # Show the user prompt
                                dlg = Dialogs.QuestionDialog(self.parent, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                            # We DO NOT want user confirmation
                            confirmations = False

                    # If the user confirms, or we skip the confirmation programatically ...
                    if result == wx.ID_YES:
                        # For each Quote in the selected items ...
                        for item in selItems:
                            # ... grab the individual item ...
                            sel = item
                            # If data is a list, there are multiple Quote nodes to paste
                            if isinstance(data, list):
                                # Iterate through the Quote nodes
                                for datum in data:
                                    # ... and paste the data
                                    DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'], confirmations=False)
                            # if data is NOT a list, it is a single Quote node to paste
                            else:
                                # ... and paste the data
                                DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'], confirmations=confirmations)
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:  # Open
            # This is functionally the same as double-clicking, although this way, we can open multiple Quotes at once!
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Quote

        elif n == 4:  # Add Quote
            try:
                # Get the Document Selection information from the ControlObject.
                (documentNum, startChar, endChar, text) = self.parent.ControlObject.GetDocumentSelectionInfo()
                # If there's a selection in the text ...
                if text != '':
                    # ... copy it to the clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                    # Add the Quote.
                    self.parent.add_quote(coll_name)
                else:
                    # Create an error message.
                    msg = _('You must make a selection in a document to be able to add a Quote.')
                    errordlg = Dialogs.InfoDialog(None, msg)
                    errordlg.ShowModal()
                    errordlg.Destroy()
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

        elif n == 5:  # Add Clip
            try:
                # Get the Transcript Selection information from the ControlObject.
                (transcriptNum, startTime, endTime, text) = self.parent.ControlObject.GetTranscriptSelectionInfo()
                # If a selection has been made ...
                if text != '':
                    # ... copy that to the Clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                # Add the Clip
                self.parent.add_clip(coll_name)
            except:
                msg = _('You must make a selection in a transcript to be able to add a Clip.')
                errordlg = Dialogs.InfoDialog(None, msg)
                errordlg.ShowModal()
                errordlg.Destroy()

        elif n == 6:  # Add Multi-transcript Clip
            # Create the appropriate Clip object
            tempClip = self.CreateMultiTranscriptClip(selData.parent, coll_name)
            # Add the Clip, using EditClip because we don't want the overhead of the Drag-and-Drop architecture
            self.parent.edit_clip(tempClip)
            # Get the Collection info
            tempCollection = Collection.Collection(tempClip.collection_num)
            # In this case, the clip has been assigned the WRONG Sort Order.  Let's wipe it out to signal that the
            # ChangeClipOrder call should reposition the clip

            tempClip.lock_record()
            tempClip.sort_order = 0
            tempClip.db_save()
            tempClip.unlock_record()
            
            # Now change the Sort Order, and if it succeeds ...
            if DragAndDropObjects.ChangeClipOrder(self, sel, tempClip, tempCollection):
                # ... let's send a message telling others they need to re-order this collection!
                if not TransanaConstants.singleUserVersion:
                    # Start by getting the Collection's Node Data
                    nodeData = tempCollection.GetNodeData()
                    # Indicate that we're sending a Collection Node
                    msg = "CollectionNode"
                    # Iterate through the nodes
                    for node in nodeData:
                        # ... add the appropriate seperator and then the node name
                        msg += ' >|< %s' % node

                    if DEBUG:
                        print 'Message to send = "OC %s"' % msg

                    # Send the Order Collection Node message
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("OC %s" % msg)

        elif n == 7:  # Add Snapshot
            # Get the Collection (the Clip's Parent)
            coll = Collection.Collection(coll_name, coll_parent_num)
            # Call the Add Snapshot dialog with the current Collection Number
            self.parent.add_snapshot(coll.number)

        elif n == 8:  # Add Quote Note
            self.parent.add_note(quoteNum=selData.recNum)

        elif n == 9:  # Merge Quotes
            
            # Load the Selected Quote.
            quote = Quote.Quote(selData.recNum)
            # There is no merging Orphan Clips!
            if quote.source_document_num == 0:
                # ... build and display an appropriate message.
                msg = unicode(_('There are no Quotes that can be merged with Quote "%s".'), 'utf8')
                dlg = Dialogs.InfoDialog(TransanaGlobal.menuWindow, msg % quote.id)
                dlg.ShowModal()
                dlg.Destroy()
                return
            # Find out if there are adjacent Quotes
            adjacentQuotes = DBInterface.FindAdjacentQuotes(quote.source_document_num, quote.start_char, quote.end_char)
            # If there are NO adjacent (mergeable) Quotes ...
            if len(adjacentQuotes) == 0:
                # ... build and display an appropriate message.
                msg = unicode(_('There are no Quotes that can be merged with Quote "%s".'), 'utf8')
                dlg = Dialogs.InfoDialog(TransanaGlobal.menuWindow, msg % quote.id)
                dlg.ShowModal()
                dlg.Destroy()
            # If there ARE adjacent (mergeable) Quotes ...
            else:
                # Initialize the quote number for the merged quote to 0
                mergedQuoteNum = 0
                # We need to lock all the mergeable quotes.  We will need to track them in a dictionary so they can later be unlocked.
                lockedQuotes = {}
                # Begin exception handling
                try:
                    # Try to get a Record Lock
                    quote.lock_record()
                    # We will need to remember the original ID of the Quote so we can tell if it changed.
                    originalQuoteID = quote.id
                # Handle the exception if the record is locked
                except RecordLockedError, e:
                    self.parent.handle_locked_record(e, _("Quote"), quote.id)
                # If the record is not locked, keep going.
                else:
                    # Let's lock all the adjacent quotes.  We can't MERGE a quote we can't lock (and therefore delete)
                    try:
                        # For all adjacent Quotes ...
                        for tmpQuoteData in adjacentQuotes:
                            # ... create a Quote object ...
                            tmpQuote = Quote.Quote(tmpQuoteData[0])
                            # ... lock that Quote object ...
                            tmpQuote.lock_record()
                            # ... and store that Quote object in the lockedQuotes dictionary so it can be unlocked later
                            lockedQuotes[tmpQuote.number] = tmpQuote
                    # If one of the adjacent Quotes cannot be locked ...
                    except RecordLockedError, e:
                        # ... Report the problem to the user ...
                        self.parent.handle_locked_record(e, _("Quote"), tmpQuote.GetNodeString(True))
                        # ... unlock the original Quote ...
                        quote.unlock_record()
                        # ... and iterate through all the locked Quotes ...
                        for quoteID in lockedQuotes.keys():
                            # ... unlocking them
                            lockedQuotes[quoteID].unlock_record()
                        # Now exit this method, as the merge has failed.
                        return

                    # Now create the Quote Properties form, passing the mergeable Quotes list to signal that a Merge is requested.
                    dlg = QuotePropertiesForm.QuotePropertiesForm(self, -1, _("Merge Quotes"), quote, mergeList=adjacentQuotes)
                    # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
                    contin = True
                    # While the "continue" flag is True ...
                    while contin:
                        # Start exception handling
                        try:
                            # Display the Quote Properties Dialog Box and get the data from the user
                            if dlg.get_input() != None:
                                # Try to save the Quote Data
                                quote.db_save()
                                # Get the source document's object if it is open somewhere in this copy of Transana
                                openSourceDocument = self.parent.ControlObject.GetOpenDocumentObject(Document.Document, quote.source_document_num)
                                # If it is open somewhere ...
                                if openSourceDocument != None:
                                    # ... then reload it to update its QuotePositions, which may have changed due to a Quote Merge.
                                    openSourceDocument.db_load_by_num(openSourceDocument.number)
                                # See if the Quote ID has been changed.  If it has, update the tree.
                                if quote.id != originalQuoteID:
                                    # Rename the Quote's Tree node
                                    self.SetItemText(sel, quote.id)
                                    # If we're in the Multi-User mode, we need to send a message about the change
                                    if not TransanaConstants.singleUserVersion:
                                        # Begin constructing the message with the old and new names for the node
                                        msg = " >|< %s >|< %s" % (originalQuoteID, quote.id)
                                        # Get the full Node Branch by climbing it to two levels above the root
                                        while (self.GetItemParent(self.GetItemParent(sel)) != self.GetRootItem()):
                                            # Update the selected node indicator
                                            sel = self.GetItemParent(sel)
                                            # Prepend the new Node's name on the Message with the appropriate seperator
                                            msg = ' >|< ' + self.GetItemText(sel) + msg
                                        # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                        # This must be untranslated to avoid problems in mixed-language environments.
                                        # Prepend these on the Messsage
                                        msg = "QuoteNode >|< Collections" + msg
                                        if DEBUG:
                                            print 'Message to send = "RN %s"' % msg
                                        # Send the Rename Node message
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                                # Now let's communicate with other Transana instances if we're in Multi-user mode
                                if not TransanaConstants.singleUserVersion:
                                    msg = 'Quote %d' % quote.number
                                    if DEBUG:
                                        print 'Message to send = "UKL %s"' % msg
                                    if TransanaGlobal.chatWindow != None:
                                        # Send the "Update Keyword List" message
                                        TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                                # If a Merge Quote has been selected in the Merge Dialog ...
                                if dlg.mergeItemIndex > -1:
                                    # ... get the merged Quote, so we can delete it.
                                    mergedQuote = Quote.Quote(dlg.mergeList[dlg.mergeItemIndex][0])
                                    # Delete the Merged Quote.  If the Merge Quote is loaded ...
                                    # If THIS Quote is currently loaded, we need to remove it from the TranscriptWindow
                                    self.parent.ControlObject.CloseOpenTranscriptWindowObject(Quote.Quote, mergedQuote.number)
                                    # Start exception handling
                                    try:
                                        # Remember the Quote number of the merged Quote to use after it is deleted
                                        mergedQuoteNum = mergedQuote.number
                                        # Get the full Node List for the Merge Quote
                                        nodeList = mergedQuote.GetNodeData(True)
                                        # Unlock the Quote selected for merging so it can be deleted!
                                        lockedQuotes[mergedQuote.number].unlock_record()
                                        # Try to delete the Merge Quote, NOT initiating a Transaction
                                        delResult = mergedQuote.db_delete(use_transactions=False)
                                        # If successful, remove the Quote Node from the Database Tree
                                        if delResult:
                                            # Call the DB Tree's delete_Node method.
                                            self.delete_Node((_('Collections'),) + nodeList, 'QuoteNode')
                                    # Handle the RecordLocked exception, which arises when records are locked!
                                    except RecordLockedError, e:
                                        # Display the Exception Message, allow "continue" flag to remain true
                                        prompt = unicode(_('You cannot delete Quote "%s".\n%s'), 'utf8')
                                        errordlg = Dialogs.ErrorDialog(None, prompt % (quote.id, e.explanation))
                                        errordlg.ShowModal()
                                        errordlg.Destroy()
                                    # Handle other exceptions
                                    except:
                                        if DEBUG:
                                            import traceback
                                            traceback.print_exc(file=sys.stdout)

                                        # Display the Exception Message, allow "continue" flag to remain true
                                        prompt = unicode("%s : %s", 'utf8')
                                        errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                                        errordlg.ShowModal()
                                        errordlg.Destroy()

                                # See if the Keyword visualization needs to be updated.
                                self.parent.ControlObject.UpdateKeywordVisualization()
                                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                                if not TransanaConstants.singleUserVersion:
                                    # We need to update the Episode Keyword Visualization
                                    if DEBUG:
                                        print 'Message to send = "UKV %s %s %s"' % ('Document', 0, quote.source_document_num)
                                        
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Document', 0, quote.source_document_num))

                                # If we do all this, we don't need to continue any more.
                                contin = False
                            # If the user pressed Cancel ...
                            else:
                                # ... then we don't need to continue any more.
                                contin = False

                        # Handle "SaveError" exception
                        except SaveError:
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
                            prompt = unicode(_("Exception %s: %s"), 'utf8')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
                            errordlg.ShowModal()
                            errordlg.Destroy()

                    # When the merge is done, destroy the Quote Properties dialog
                    dlg.Destroy()

                    # Unlock the record regardless of what happens
                    quote.unlock_record()

                # Finally, iterate through the dictionary of locked Quotes ...
                for quoteID in lockedQuotes.keys():
                    # ... and if we don't have the quote that got merged (which has already been unlocked and deleted) ...
                    if quoteID != mergedQuoteNum:
                        # ... unlock the quote
                        lockedQuotes[quoteID].unlock_record()

        elif n == 10:  # Delete Quotes
            # If the Notes Browser is showing in single-user modes ...
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Quote Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # For each Quote in the selected items ...
                for item in selItems:
                    self.DeleteQuoteClipSnapshotItem(item)
                    
##                    # ... grab the individual item ...
##                    sel = item
##                    # ... and get the item's data
##                    selData = self.GetPyData(sel)
##
##                    if self.GetPyData(item).nodetype == 'QuoteNode':
##                        # Load the Selected Quote.  We DO NOT need the Quote Text here
##                        tmpObj = Quote.Quote(num=selData.recNum, skipText=True)
##                        # Get the Object Parent number (so it will persist after the Object is deleted)
##                        tmpObjParentNum = tmpObj.source_document_num
##                        tmpObjType = Quote.Quote
##                        tmpObjTypeText = 'Quote'
##                        # Get user confirmation of the Quote Delete request
##                        prompt = unicode(_('Are you sure you want to delete Quote "%s" and all related Notes?'), 'utf8')
##                        delPrompt = unicode(_('You cannot delete Quote "%s".\n%s'), 'utf8')
##                    elif self.GetPyData(item).nodetype == 'ClipNode':
##                        # Load the Selected Clip.  We DO NOT need the Clip Text here
##                        tmpObj = Clip.Clip(selData.recNum, skipText=True)
##                        # Get the Object Parent number (so it will persist after the Object is deleted)
##                        tmpObjParentNum = tmpObj.episode_num
##                        tmpObjType = Clip.Clip
##                        tmpObjTypeText = 'Clip'
##                        # Get user confirmation of the Clip Delete request
##                        prompt = unicode(_('Are you sure you want to delete Clip "%s" and all related Notes?'), 'utf8')
##                        delPrompt = unicode(_('You cannot delete Clip "%s".\n%s'), 'utf8')
##                    elif self.GetPyData(item).nodetype == 'SnapshotNode':
##                        # Load the Selected Snapshot.
##                        tmpObj = Snapshot.Snapshot(selData.recNum)
##                        # There is no Parent Object number
##                        tmpObjParentNum = 0
##                        tmpObjType = Snapshot.Snapshot
##                        tmpObjTypeText = 'Snapshot'
##                        # Get user confirmation of the Snapshot Delete request
##                        prompt = unicode(_('Are you sure you want to delete Snapshot "%s" and all related Notes?'), 'utf8')
##                        delPrompt = unicode(_('You cannot delete Snapshot "%s".\n%s'), 'utf8')
##                            
##                    dlg = Dialogs.QuestionDialog(self, prompt % self.GetItemText(sel))
##                    result = dlg.LocalShowModal()
##                    dlg.Destroy()
##
##                    # If the user confirms the Delete Request...
##                    if result == wx.ID_YES:
##                        # If THIS Object is currently loaded, we need to remove it from the TranscriptWindow
##                        self.parent.ControlObject.CloseOpenTranscriptWindowObject(tmpObjType, tmpObj.number)
##                        # Start exception handling
##                        try:
##                            # Get the Object number (so it will persist after the Object is deleted)
##                            tmpObjNum = tmpObj.number
##                            # Try to delete the Object, initiating a Transaction
##                            delResult = tmpObj.db_delete(1)
##                            # If successful, remove the Object Node from the Database Tree
##                            if delResult:
##                                if isinstance(tmpObj, Quote.Quote):
##                                    # Remove the Object's Position Data from the Source Document, if it's open
##                                    self.parent.ControlObject.RemoveQuoteFromOpenDocument(tmpObjNum, tmpObjParentNum)
##                                # Get a temporary Selection Pointer
##                                tempSel = sel
##                                # Get the full Node Branch by climbing it to one level above the root
##                                nodeList = (self.GetItemText(tempSel),)
##                                while (self.GetItemParent(tempSel) != self.GetRootItem()):
##                                    tempSel = self.GetItemParent(tempSel)
##                                    nodeList = (self.GetItemText(tempSel),) + nodeList
##
##                                # Call the DB Tree's delete_Node method.
##                                self.delete_Node(nodeList, self.GetPyData(item).nodetype)
##                                # ... and if the Notes Browser is open, ...
##                                if self.parent.ControlObject.NotesBrowserWindow != None:
##                                    # ... we need to CHECK to see if any notes were deleted.
##                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
##                                    
##                                # If the current main object is a Document and it's the Document that contains the
##                                # deleted Object, we need to update the Document Keyword Visualization!
##                                if (isinstance(self.parent.ControlObject.currentObj, Document.Document)) and \
##                                   (sourceDocNum == self.parent.ControlObject.currentObj.number):
##                                    self.parent.ControlObject.UpdateKeywordVisualization()
##                                # Even if this computer doesn't need to update the keyword visualization others, might need to.
##                                if not TransanaConstants.singleUserVersion:
##                                    # We need to pass the type of the current object, the deleted Object's record number, and
##                                    # the deleted Object's Parent number.
##                                    if DEBUG:
##                                        print 'Message to send = "UKV %s %s %s"' % (tmpObjTypeText, tmpObjNum, tmpObjParentNum)
##                                        
##                                    if TransanaGlobal.chatWindow != None:
##                                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % (tmpObjTypeText, tmpObjNum, tmpObjParentNum))
##
##                        # Handle the RecordLocked exception, which arises when records are locked!
##                        except RecordLockedError, e:
##                            # Display the Exception Message, allow "continue" flag to remain true
##                            errordlg = Dialogs.ErrorDialog(None, delPrompt % (tmpObj.id, e.explanation))
##                            errordlg.ShowModal()
##                            errordlg.Destroy()
##                        # Handle other exceptions
##                        except:
##                            if DEBUG:
##                                import traceback
##                                traceback.print_exc(file=sys.stdout)
##
##                            # Display the Exception Message, allow "continue" flag to remain true
##                            prompt = "%s : %s"
##                            if 'unicode' in wx.PlatformInfo:
##                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                                prompt = unicode(prompt, 'utf8')
##                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
##                            errordlg.ShowModal()
##                            errordlg.Destroy()

        elif n == 11:  # Locate Quote in Document
            self.parent.ControlObject.LocateQuoteInDocument(selData.recNum)

        elif n == 12:  # Insert Quote Hyperlink
            # If the current transcript is in Read Only mode ...
            if self.parent.ControlObject.ActiveTranscriptReadOnly():
                # ... inform the user
                msg = _("The current document is not editable.  The requested hyperlink cannot be inserted into the document.")
                msg += '\n\n' + _("To insert the hyperlink into the document, press the Edit Mode button on the Document Toolbar to make the document editable.")
                dlg = Dialogs.InfoDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
            else:
                # Load the TEMP file into Transcript
                self.parent.ControlObject.TranscriptInsertHyperlink("Quote", selData.recNum)

        elif n == 13:  # Quote Properties
            # Load the Quote Object
            quote = Quote.Quote(selData.recNum)
            # Edit the Quote Properties
            self.parent.edit_quote(quote)

    def OnClipCommand(self, evt):
        """Handle selections for the Clip menu."""
        n = evt.GetId() - self.cmd_id_start['ClipNode']
        # If we're in the Standard version, we need to adjust the menu numbers
        # for Add Quote (4) Add Multi-transcript Clip (6) and Add Snapshot (7)
        if not TransanaConstants.proVersion and (n >= 4):
            n += 1
        if not TransanaConstants.proVersion and (n >= 6):
            n += 2
        # If we're in the Standard version, we need to adjust the menu numbers
        # for Export Clip Video (12) and Insert Clip Hyperlink (13)
        if not TransanaConstants.proVersion and (n >= 12):
            n += 2
        
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            clip_name = self.GetItemText(sel)
            coll_name = self.GetItemText(self.GetItemParent(sel))
            coll_parent_num = self.GetPyData(self.GetItemParent(sel)).parent
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:    # Paste
            # specify the data formats to accept.
            #   Our data could be a DataTreeDragData object if the source is the Database Tree
            dfNode = wx.CustomDataFormat('DataTreeDragData')
            #   Our data could be a ClipDragDropData object if the source is the Transcript (clip creation)
            dfClip = wx.CustomDataFormat('ClipDragDropData')

            # Specify the data object to accept data for these formats
            #   A DataTreeDragData object will populate the cdoNode object
            cdoNode = wx.CustomDataObject(dfNode)
            #   A ClipDragDropData object will populate the cdoClip object
            cdoClip = wx.CustomDataObject(dfClip)

            # Create a composite Data Object from the object types defined above
            cdo = wx.DataObjectComposite()
            cdo.Add(cdoNode)
            cdo.Add(cdoClip)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get data from the Clipboard
            success = wx.TheClipboard.GetData(cdo)
            # If the data in the clipboard is in an appropriate format ...
            if success:
                # ... unPickle the data so it's in a usable form
                # First, let's try to get the DataTreeDragData object
                try:
                    data = cPickle.loads(cdoNode.GetData())
                except:
                    data = None
                    # If this fails, that's okay
                    pass

                # Let's also try to get the ClipDragDropData object
                try:
                    data2 = cPickle.loads(cdoClip.GetData())
                except:
                    data2 = None
                    # If this fails, that's okay
                    pass

                # if we didn't get the DataTreeDragData object, we need to substitute the ClipDragDropData object
                if data == None:
                    data = data2
                    
                # If our clipboard data is text for creating a clip ...
                if type(data) == type(DragAndDropObjects.ClipDragDropData()):
                    # ... then create a clip!  (I don't think this ever gets called!??)
                    DragAndDropObjects.CreateClip(data, selData, self, sel)
                # If we have drag data fron the Database Tree ...
                else:
                    # Multiple SOURCE items
                    if isinstance(data, list):
                        # One DESTINATION item
                        if len(selItems) == 1:
                            # If the SOURCE is a list of Keywords, get User confirmation
                            if data[0].nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                                    data1 = unicode(_('Clip'), 'utf8')
                                else:
                                    prompt = _('Do you want to add multiple Keywords to %s "%s"?')
                                    data1 = _('Clip')
                                # Set up data to go with the prompt
                                promptdata = (data1, self.GetItemText(sel))
                                # Show the user prompt
                                dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                        # Multiple DESTINATION items
                        else:
                            # If the SOURCE is a list of Keywords, get User confirmation
                            if data[0].nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add multiple Keywords to multiple %s?'), 'utf8')
                                    data1 = unicode(_('Clips'), 'utf8')
                                else:
                                    prompt = _('Do you want to add multiple Keywords to multiple %s?')
                                    data1 = _('Clips')
                                # Set up data to go with the prompt
                                promptdata = (data1,)
                                # Set up data to go with the prompt
                                dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                    # One SOURCE item
                    else:
                        # One DESTINATION item
                        if len(selItems) == 1:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                            # Prompt gets handled by DragAndDropObjects.DropKeyword().  We DO want user confirmation
                            confirmations = True
                        # Multiple DESTINATION items
                        else:
                            # If the SOURCE is a Keyword, get User confirmation
                            if data.nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add Keyword "%s:%s" to multiple %s?'), 'utf8') 
                                    data1 = unicode(_('Clips'), 'utf8')
                                else:
                                    prompt = _('Do you want to add Keyword "%s:%s" to multiple %s?')
                                    data1 = _('Clips')
                                # Set up data to go with the prompt
                                promptdata = (data.parent, data.text, data1)
                                # Show the user prompt
                                dlg = Dialogs.QuestionDialog(self.parent, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                            # We DO NOT want user confirmation
                            confirmations = False

                    # If the user confirms, or we skip the confirmation programatically ...
                    if result == wx.ID_YES:
                        # For each Clip in the selected items ...
                        for item in selItems:
                            # ... grab the individual item ...
                            sel = item
                            # If data is a list, there are multiple Clip nodes to paste
                            if isinstance(data, list):
                                # Iterate through the Clip nodes
                                for datum in data:
                                    # ... and paste the data
                                    DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'], confirmations=False)
                            # if data is NOT a list, it is a single Clip node to paste
                            else:
                                # ... and paste the data
                                DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'], confirmations=confirmations)
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:    # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Clip

        elif n == 4:    # Add Quote
            try:
                # Get the Document Selection information from the ControlObject.
                (documentNum, startChar, endChar, text) = self.parent.ControlObject.GetDocumentSelectionInfo()
                # If there's a selection in the text ...
                if text != '':
                    # ... copy it to the clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                    # Add the Quote.
                    self.parent.add_quote(coll_name)
                else:
                    # Create an error message.
                    msg = _('You must make a selection in a document to be able to add a Quote.')
                    errordlg = Dialogs.InfoDialog(None, msg)
                    errordlg.ShowModal()
                    errordlg.Destroy()
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

        elif n == 5:    # Add Clip
            try:
                # Get the Transcript Selection information from the ControlObject.
                (transcriptNum, startTime, endTime, text) = self.parent.ControlObject.GetTranscriptSelectionInfo()
                # If a selection has been made ...
                if text != '':
                    # ... copy that to the Clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                # Add the Clip
                self.parent.add_clip(coll_name)
            except:
                msg = _('You must make a selection in a transcript to be able to add a Clip.')
                errordlg = Dialogs.InfoDialog(None, msg)
                errordlg.ShowModal()
                errordlg.Destroy()

        elif n == 6:    # Add Multi-transcript Clip
            # Create the appropriate Clip object
            tempClip = self.CreateMultiTranscriptClip(selData.parent, coll_name)
            # Add the Clip, using EditClip because we don't want the overhead of the Drag-and-Drop architecture
            self.parent.edit_clip(tempClip)
            # Get the Collection info
            tempCollection = Collection.Collection(tempClip.collection_num)
            # In this case, the clip has been assigned the WRONG Sort Order.  Let's wipe it out to signal that the
            # ChangeClipOrder call should reposition the clip

            tempClip.lock_record()
            tempClip.sort_order = 0
            tempClip.db_save()
            tempClip.unlock_record()
            
            # Now change the Sort Order, and if it succeeds ...
            if DragAndDropObjects.ChangeClipOrder(self, sel, tempClip, tempCollection):
                # ... let's send a message telling others they need to re-order this collection!
                if not TransanaConstants.singleUserVersion:
                    # Start by getting the Collection's Node Data
                    nodeData = tempCollection.GetNodeData()
                    # Indicate that we're sending a Collection Node
                    msg = "CollectionNode"
                    # Iterate through the nodes
                    for node in nodeData:
                        # ... add the appropriate seperator and then the node name
                        msg += ' >|< %s' % node

                    if DEBUG:
                        print 'Message to send = "OC %s"' % msg

                    # Send the Order Collection Node message
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("OC %s" % msg)

        elif n == 7:    # Add Snapshot
            # Get the Collection (the Clip's Parent)
            coll = Collection.Collection(coll_name, coll_parent_num)
            # Call the Add Snapshot dialog with the current Collection Number
            self.parent.add_snapshot(coll.number)

        elif n == 8:    # Add Note
            self.parent.add_note(clipNum=selData.recNum)

        elif n == 9:    # Merge Clips
            # Load the Selected Clip.  We DO need the Transcript Text
            clip = Clip.Clip(selData.recNum)
            # We need a list of the Clip Source Transcript information (number, clip_start, clip_stop).  Initialize that here.
            trInfo = []
            # now iterate through the clip transcripts and add their source numbers, start, and end times to the list.
            for tr in clip.transcripts:
                trInfo.append((tr.source_transcript, tr.clip_start, tr.clip_stop))
            # We also need a list of the Clip Media Files and Audio Inclusion information.  Initialize that
            # with the primary media file and audio information
            trFiles = [(clip.media_filename, clip.audio)]
            # Iterate through the additional media files ...
            for addFile in clip.additional_media_files:
                # ... adding media filename and audio information.
                trFiles.append((addFile['filename'], addFile['audio']))
            # Find out if there are adjacent clips
            adjacentClips = DBInterface.FindAdjacentClips(clip.episode_num, clip.clip_start, clip.clip_stop, trInfo, trFiles)
            # If there are NO adjacent (mergeable) clips ...
            if len(adjacentClips) == 0:
                # ... build and display an appropriate message.
                msg = _('There are no clips that can be merged with clip "%s".')
                if 'unicode' in wx.PlatformInfo:
                    msg = unicode(msg, 'utf8')
                dlg = Dialogs.InfoDialog(TransanaGlobal.menuWindow, msg % clip.id)
                dlg.ShowModal()
                dlg.Destroy()
            # If there ARE adjacent (mergeable) clips ...
            else:
                # Initialize the clip number for the merged clip to 0
                mergedClipNum = 0
                # We need to lock all the mergeable clips.  We will need to track them in a dictionary so they can later be unlocked.
                lockedClips = {}
                # Begin exception handling
                try:
                    # Try to get a Record Lock
                    clip.lock_record()
                    # We will need to remember the original ID of clip so we can tell if it changed.
                    originalClipID = clip.id
                # Handle the exception if the record is locked
                except RecordLockedError, e:
                    self.parent.handle_locked_record(e, _("Clip"), clip.id)
                # If the record is not locked, keep going.
                else:
                    # Let's lock all the adjacent clips.  We can't MERGE a clip we can't lock (and therefore delete)
                    try:
                        # For all adjacent clips ...
                        for tmpClipData in adjacentClips:
                            # ... create a clip object ...  (We CANNOT speed the load by not loading the Clip Transcript(s),
                            #                                as we need to lock Clip Transcripts!)
                            tmpClip = Clip.Clip(tmpClipData[0])
                            # ... lock that clip object ...
                            tmpClip.lock_record()
                            # ... and store that clip object in the lockedClips dictionary so it can be unlocked later
                            lockedClips[tmpClip.number] = tmpClip
                    # If one of the adjacent clips cannot be locked ...
                    except RecordLockedError, e:
                        # ... Report the problem to the user ...
                        self.parent.handle_locked_record(e, _("Clip"), tmpClip.GetNodeString(True))
                        # ... unlock the original clip ...
                        clip.unlock_record()
                        # ... and iterate through all the locked clips ...
                        for clipID in lockedClips.keys():
                            # ... unlocking them
                            lockedClips[clipID].unlock_record()
                        # Now exit this method, as the merge has failed.
                        return
                        
                    # Now create the Clip Properties form, passing the mergeable clips list to signal that a Merge is requested.
                    dlg = ClipPropertiesForm.ClipPropertiesForm(self, -1, _("Merge Clips"), clip, mergeList=adjacentClips)
                    # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
                    contin = True
                    # While the "continue" flag is True ...
                    while contin:
                        # Start exception handling
                        try:
                            # Display the Clip Properties Dialog Box and get the data from the user
                            if dlg.get_input() != None:
                                # Try to save the Clip Data
                                clip.db_save()

                                # If any keywords that served as Keyword Examples that got removed from the Clip,
                                # we need to remove them from the Database Tree.  
                                for (keywordGroup, keyword, clipNum) in dlg.keywordExamplesToDelete:
                                    # Prepare the Node List for removing the Keyword Example Node, using the clip's original name
                                    nodeList = (_('Keywords'), keywordGroup, keyword, originalClipID)
                                    # Call the DB Tree's delete_Node method.  Include the Clip Record Number so the correct Clip entry will be removed.
                                    self.delete_Node(nodeList, 'KeywordExampleNode', clip.number)

                                # See if the Clip ID has been changed.  If it has, update the tree.
                                if clip.id != originalClipID:

                                    # Look for Keyword Examples in this clip.
                                    for (kwg, kw, clipNumber, clipID) in DBInterface.list_all_keyword_examples_for_a_clip(clip.number):
                                        # Get the current Node path for the KWE
                                        nodeList = (_('Keywords'), kwg, kw, originalClipID)
                                        # Get the Keyword Example node
                                        exampleNode = self.select_Node(nodeList, 'KeywordExampleNode')
                                        # exampleNode will be None if node could not be found (which CAN happen because of keyword examples in the merged clip) 
                                        if exampleNode != None:
                                            # Rename the node
                                            self.SetItemText(exampleNode, clip.id)
                                            # If we're in the Multi-User mode, we need to send a message about the change
                                            if not TransanaConstants.singleUserVersion:
                                                # Begin constructing the message with the old and new names for the node
                                                msg = " >|< %s >|< %s" % (originalClipID, clip.id)
                                                # Get the full Node Branch by climbing it to two levels above the root
                                                while (self.GetItemParent(self.GetItemParent(exampleNode)) != self.GetRootItem()):
                                                    # Update the selected node indicator
                                                    exampleNode = self.GetItemParent(exampleNode)
                                                    # Prepend the new Node's name on the Message with the appropriate seperator
                                                    msg = ' >|< ' + self.GetItemText(exampleNode) + msg
                                                # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                                # This must be untranslated to avoid problems in mixed-language environments.
                                                # Prepend these on the Messsage
                                                msg = "KeywordExampleNode >|< Keywords" + msg
                                                if DEBUG:
                                                    print 'Message to send = "RN %s"' % msg
                                                # Send the Rename Node message
                                                if TransanaGlobal.chatWindow != None:
                                                    TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                                    # Rename the Clip's Tree node
                                    self.SetItemText(sel, clip.id)
                                    # If we're in the Multi-User mode, we need to send a message about the change
                                    if not TransanaConstants.singleUserVersion:
                                        # Begin constructing the message with the old and new names for the node
                                        msg = " >|< %s >|< %s" % (originalClipID, clip.id)
                                        # Get the full Node Branch by climbing it to two levels above the root
                                        while (self.GetItemParent(self.GetItemParent(sel)) != self.GetRootItem()):
                                            # Update the selected node indicator
                                            sel = self.GetItemParent(sel)
                                            # Prepend the new Node's name on the Message with the appropriate seperator
                                            msg = ' >|< ' + self.GetItemText(sel) + msg
                                        # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                        # This must be untranslated to avoid problems in mixed-language environments.
                                        # Prepend these on the Messsage
                                        msg = "ClipNode >|< Collections" + msg
                                        if DEBUG:
                                            print 'Message to send = "RN %s"' % msg
                                        # Send the Rename Node message
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                                # Now let's communicate with other Transana instances if we're in Multi-user mode
                                if not TransanaConstants.singleUserVersion:
                                    msg = 'Clip %d' % clip.number
                                    if DEBUG:
                                        print 'Message to send = "UKL %s"' % msg
                                    if TransanaGlobal.chatWindow != None:
                                        # Send the "Update Keyword List" message
                                        TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                                # If a Merge Clip has been selected in the Merge Dialog ...
                                if dlg.mergeItemIndex > -1:
                                    # ... get the merged Clip, so we can delete it.  Don't skipText!
                                    mergedClip = Clip.Clip(dlg.mergeList[dlg.mergeItemIndex][0])
                                    # If the merged clip contains KW Examples, we need to rename the example nodes.
                                    # Iterate through the list of keywords
                                    for kw in mergedClip.keyword_list:
                                        # If the keyword is an example ...
                                        if kw.example:
                                            # It's possible that the example keyword was removed from the form.  We need to know.
                                            kwFound = False
                                            # Iterate through the original clip's Keyword List ...
                                            for kw2 in clip.keyword_list:
                                                # ... Look for matching keyword values ...
                                                if (kw.keywordGroup == kw2.keywordGroup) and \
                                                   (kw.keyword == kw2.keyword):
                                                    # ... If found, mark the clip's keyword as an example.
                                                    kw2.example = 1
                                                    # Save the original clip (again) so make this permanent.
                                                    clip.db_save()
                                                    # We need to update the Keyword Example node in the database tree.
                                                    # Renaming the node isn't sufficient here, as the node's data wouldn't get updated.
                                                    # Therefore, we delete the OLD node and add the NEW node.
                                                    self.delete_Node((_('Keywords'), kw2.keywordGroup, kw2.keyword, mergedClip.id), 'KeywordExampleNode', exampleClipNum=mergedClip.number, sendMessage=True)
                                                    self.add_Node('KeywordExampleNode', (_('Keywords'), kw2.keywordGroup, kw2.keyword, clip.id), clip.number, clip.collection_num, expandNode=False)
                                                    # If Multi-User ...
                                                    if TransanaGlobal.chatWindow != None:
                                                        # ... we need to signal that a keyword example has been added.
                                                        # (The multi-user KWE delete messaging is handles elsewhere, but I don't know where.)
                                                        TransanaGlobal.chatWindow.SendMessage("AKE %d >|< %s >|< %s >|< %s" % (clip.number, kw2.keywordGroup, kw2.keyword, clip.id))
                                                    # Finally, note that the keyword was found and updated
                                                    kwFound = True
                                            # If the appropriate keyword was never found, this means that the merged clip keyword
                                            # example was removed during clip merge.  If this is the case ...
                                            if not kwFound:
                                                # ... we need to delete the KWE record from the database tree.
                                                # (The multi-user KWE delete messaging is handles elsewhere, but I don't know where.)
                                                self.delete_Node((_('Keywords'), kw.keywordGroup, kw.keyword, mergedClip.id), 'KeywordExampleNode', exampleClipNum=mergedClip.number, sendMessage=True)

                                    # Delete the Merged Clip.  If the Merge Clip is loaded ...
                                    if isinstance(self.parent.ControlObject.currentObj, Clip.Clip) and (self.parent.ControlObject.currentObj.number == mergedClip.number):
                                        # ... Start by clearing all current objects
                                        self.parent.ControlObject.ClearAllWindows()
                                    # Start exception handling
                                    try:
                                        # Remember the clip number of the merged clip to use after it is deleted
                                        mergedClipNum = mergedClip.number
                                        # Get the full Node List for the Merge Clip
                                        nodeList = mergedClip.GetNodeData(True)
                                        # Unlock the clip selected for merging so it can be deleted!
                                        lockedClips[mergedClip.number].unlock_record()
                                        # Try to delete the Merge Clip, NOT initiating a Transaction
                                        delResult = mergedClip.db_delete(use_transactions=False, examplesPrompt=False)
                                        # If successful, remove the Clip Node from the Database Tree
                                        if delResult:
                                            # Call the DB Tree's delete_Node method.
                                            self.delete_Node((_('Collections'),) + nodeList, 'ClipNode')
                                    # Handle the RecordLocked exception, which arises when records are locked!
                                    except RecordLockedError, e:
                                        # Display the Exception Message, allow "continue" flag to remain true
                                        if 'unicode' in wx.PlatformInfo:
                                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                            prompt = unicode(_('You cannot delete Clip "%s".\n%s'), 'utf8')
                                        else:
                                            prompt = _('You cannot delete Clip "%s".\n%s')
                                        errordlg = Dialogs.ErrorDialog(None, prompt % (clip.id, e.explanation))
                                        errordlg.ShowModal()
                                        errordlg.Destroy()
                                    # Handle other exceptions
                                    except:
                                        if DEBUG:
                                            import traceback
                                            traceback.print_exc(file=sys.stdout)

                                        # Display the Exception Message, allow "continue" flag to remain true
                                        prompt = "%s : %s"
                                        if 'unicode' in wx.PlatformInfo:
                                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                            prompt = unicode(prompt, 'utf8')
                                        errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                                        errordlg.ShowModal()
                                        errordlg.Destroy()

                                # See if the Keyword visualization needs to be updated.
                                self.parent.ControlObject.UpdateKeywordVisualization()
                                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                                if not TransanaConstants.singleUserVersion:
                                    # We need to update the Episode Keyword Visualization
                                    if DEBUG:
                                        print 'Message to send = "UKV %s %s %s"' % ('Episode', clip.episode_num, 0)
                                        
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', clip.episode_num, 0))

                                # If we do all this, we don't need to continue any more.
                                contin = False
                            # If the user pressed Cancel ...
                            else:
                                # ... then we don't need to continue any more.
                                contin = False


                        # Handle "SaveError" exception
                        except SaveError:
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

                    # When the merge is done, destroy the Clip Properties dialog
                    dlg.Destroy()

                    # Unlock the record regardless of what happens
                    clip.unlock_record()

                # Finally, iterate through the dictionary of locked clips ...
                for clipID in lockedClips.keys():
                    # ... and if we don't have the clip that got merged (which has already been unlocked and deleted) ...
                    if clipID != mergedClipNum:
                        # ... unlock the clip
                        lockedClips[clipID].unlock_record()

        elif n == 10:    # Delete Clip
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Clip Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # For each Clip in the selected items ...
                for item in selItems:
                    self.DeleteQuoteClipSnapshotItem(item)
                    
##                    # ... grab the individual item ...
##                    sel = item
##                    # ... and get the item's data
##                    selData = self.GetPyData(sel)
##
##                    # Load the Selected Clip.  We DO need the Clip Transcript(s) here
##                    clip = Clip.Clip(selData.recNum)
##                    # Remember the Clip Number, to use after the clip has been deleted
##                    clipNum = clip.number
##                    # Remember the original clip's Episode Number for use later, after the clip has been deleted
##                    clipEpisodeNum = clip.episode_num
##                    # Get user confirmation of the Clip Delete request
##                    if 'unicode' in wx.PlatformInfo:
##                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                        prompt = unicode(_('Are you sure you want to delete Clip "%s" and all related Notes?'), 'utf8')
##                    else:
##                        prompt = _('Are you sure you want to delete Clip "%s" and all related Notes?')
##                    dlg = Dialogs.QuestionDialog(self, prompt % self.GetItemText(sel))
##                    result = dlg.LocalShowModal()
##                    dlg.Destroy()
##
##                    # If the user confirms the Delete Request...
##                    if result == wx.ID_YES:
###                        # If THIS Clip is loaded ...
###                        if isinstance(self.parent.ControlObject.currentObj, Clip.Clip) and (self.parent.ControlObject.currentObj.number == clip.number):
###                            # ... Start by clearing all current objects
###                            self.parent.ControlObject.ClearAllWindows()
##
##                        # Determine what Keyword Examples exist for the specified Clip so that they can be removed from the
##                        # Database Tree if the delete succeeds.  We must do that first, as the Clips themselves and the
##                        # ClipKeywords Records will be deleted later!
##                        kwExamples = DBInterface.list_all_keyword_examples_for_a_clip(selData.recNum)
##
##                        try:
##                            # Try to delete the Clip, initiating a Transaction
##                            delResult = clip.db_delete(1)
##                            # If successful, remove the Clip Node from the Database Tree
##                            if delResult:
##                                # Get a temporary Selection Pointer
##                                tempSel = sel
##                                # Get the full Node Branch by climbing it to one level above the root
##                                nodeList = (self.GetItemText(tempSel),)
##                                while (self.GetItemParent(tempSel) != self.GetRootItem()):
##                                    tempSel = self.GetItemParent(tempSel)
##                                    nodeList = (self.GetItemText(tempSel),) + nodeList
##
##                                # Deleting all these ClipKeyword records needs to remove Keyword Example Nodes in the DBTree.
##                                # That needs to be done here in the User Interface rather than in the Clip Object, as that is
##                                # a user interface issue.  The Clip Record and the Clip Keywords Records get deleted, but
##                                # the user interface does not get cleaned up by deleting the Clip Object.
##                                for (kwg, kw, clipNum, clipID) in kwExamples:
##                                    self.delete_Node((_("Keywords"), kwg, kw, clipID), 'KeywordExampleNode', exampleClipNum = clipNum)
##
##                                # Call the DB Tree's delete_Node method.
##                                self.delete_Node(nodeList, 'ClipNode')
##                                # ... and if the Notes Browser is open, ...
##                                if self.parent.ControlObject.NotesBrowserWindow != None:
##                                    # ... we need to CHECK to see if any notes were deleted.
##                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
##                                # If this clip is from the episode which is currently being displayed ...
##
##                                # If the current main object is an Episode and it's the episode that contains the
##                                # deleted Clip, we need to update the Keyword Visualization!
##                                if (isinstance(self.parent.ControlObject.currentObj, Episode.Episode)) and \
##                                   (clipEpisodeNum == self.parent.ControlObject.currentObj.number):
##                                    self.parent.ControlObject.UpdateKeywordVisualization()
##                                # Even if this computer doesn't need to update the keyword visualization others, might need to.
##                                if not TransanaConstants.singleUserVersion:
##                                    # We need to pass the type of the current object, the deleted Clip's record number, and
##                                    # the deleted Clip's Episode number.
##                                    if DEBUG:
##                                        print 'Message to send = "UKV %s %s %s"' % ('Clip', clipNum, clipEpisodeNum)
##                                        
##                                    if TransanaGlobal.chatWindow != None:
##                                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', clipNum, clipEpisodeNum))
##                        # Handle the RecordLocked exception, which arises when records are locked!
##                        except RecordLockedError, e:
##                            # Display the Exception Message, allow "continue" flag to remain true
##                            if 'unicode' in wx.PlatformInfo:
##                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                                prompt = unicode(_('You cannot delete Clip "%s".\n%s'), 'utf8')
##                            else:
##                                prompt = _('You cannot delete Clip "%s".\n%s')
##                            errordlg = Dialogs.ErrorDialog(None, prompt % (clip.id, e.explanation))
##                            errordlg.ShowModal()
##                            errordlg.Destroy()
##                        # Handle other exceptions
##                        except:
##                            if DEBUG:
##                                import traceback
##                                traceback.print_exc(file=sys.stdout)
##
##                            # Display the Exception Message, allow "continue" flag to remain true
##                            prompt = "%s : %s"
##                            if 'unicode' in wx.PlatformInfo:
##                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                                prompt = unicode(prompt, 'utf8')
##                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
##                            errordlg.ShowModal()
##                            errordlg.Destroy()

        elif n == 11:    # Locate Clip in Episode

            self.parent.ControlObject.LocateClipInEpisode(selData.recNum)

        elif n == 12:    # Export Clip Video
            # Load the Clip.  We can speed the load by not loading the Clip Transcript(s).
            clip = Clip.Clip(selData.recNum, skipText=True)
            # Create the Media Conversion dialog, including Clip Information so we export only the clip segment
            convertDlg = MediaConvert.MediaConvert(self, clip.media_filename, clip.clip_start - clip.offset, clip.clip_stop - clip.clip_start, clip.id)
            # Show the Media Conversion Dialog
            convertDlg.ShowModal()
            # Destroy the Media Conversion Dialog
            convertDlg.Destroy()
            
        elif n == 13:  # Insert Clip Hyperlink
            # If the current transcript is in Read Only mode ...
            if self.parent.ControlObject.ActiveTranscriptReadOnly():
                # ... inform the user
                msg = _("The current document is not editable.  The requested hyperlink cannot be inserted into the document.")
                msg += '\n\n' + _("To insert the hyperlink into the document, press the Edit Mode button on the Document Toolbar to make the document editable.")
                dlg = Dialogs.InfoDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
            else:
                # Load the TEMP file into Transcript
                self.parent.ControlObject.TranscriptInsertHyperlink("Clip", selData.recNum)

        elif n == 14:    # Clip Properties
            # Load the Clip.  We DO need the Clip Transcript(s) here.
            clip = Clip.Clip(selData.recNum)
            # Edit the Clip properties
            self.parent.edit_clip(clip)
##            # If the current main object is an Episode and it's the episode that contains the
##            # edited Clip, we need to update the Keyword Visualization!
##            if (isinstance(self.parent.ControlObject.currentObj, Episode.Episode)) and \
##               (clip.episode_num == self.parent.ControlObject.currentObj.number):
##                self.parent.ControlObject.UpdateKeywordVisualization()
##            # Even if this computer doesn't need to update the keyword visualization others, might need to.
##            if not TransanaConstants.singleUserVersion:
##                # We need to pass the type of the current object, the Clip's record number, and
##                # the Clip's Episode number.
##                if DEBUG:
##                    print 'Message to send = "UKV %s %s %s"' % ('Clip', clip.number, clip.episode_num)
##                    
##                if TransanaGlobal.chatWindow != None:
##                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', clip.number, clip.episode_num))

        else:
            raise MenuIDError

    def OnSnapshotCommand(self, evt):
        """Handle selections for the Snapshot menu."""
        n = evt.GetId() - self.cmd_id_start["SnapshotNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
#            snapshot_name = self.GetItemText(sel)
            coll_name = self.GetItemText(self.GetItemParent(sel))
            coll_parent_num = self.GetPyData(self.GetItemParent(sel)).parent

        # Get the list of selected items
        selItems = self.GetSelections()
        if n == 0:        # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:      # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:      # Paste
            # specify the data formats to accept.
            #   Our data could be a DataTreeDragData object if the source is the Database Tree
            dfNode = wx.CustomDataFormat('DataTreeDragData')
            #   Our data could be a ClipDragDropData object if the source is the Transcript (clip creation)
            dfClip = wx.CustomDataFormat('ClipDragDropData')

            # Specify the data object to accept data for these formats
            #   A DataTreeDragData object will populate the cdoNode object
            cdoNode = wx.CustomDataObject(dfNode)
            #   A ClipDragDropData object will populate the cdoClip object
            cdoClip = wx.CustomDataObject(dfClip)

            # Create a composite Data Object from the object types defined above
            cdo = wx.DataObjectComposite()
            cdo.Add(cdoNode)
            cdo.Add(cdoClip)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get data from the Clipboard
            success = wx.TheClipboard.GetData(cdo)
            # If the data in the clipboard is in an appropriate format ...
            if success:
                # ... unPickle the data so it's in a usable form
                # First, let's try to get the DataTreeDragData object
                try:
                    data = cPickle.loads(cdoNode.GetData())
                except:
                    data = None
                    # If this fails, that's okay
                    pass

                # Let's also try to get the ClipDragDropData object
                try:
                    data2 = cPickle.loads(cdoClip.GetData())
                except:
                    data2 = None
                    # If this fails, that's okay
                    pass

                # if we didn't get the DataTreeDragData object, we need to substitute the ClipDragDropData object
                if data == None:
                    data = data2
                    
                # If our clipboard data is text for creating a clip ...
                if type(data) == type(DragAndDropObjects.ClipDragDropData()):
                    # ... then create a clip!  (I don't think this ever gets called!??)
                    DragAndDropObjects.CreateClip(data, selData, self, sel)
                # If we have drag data fron the Database Tree ...
                else:
                    # Multiple SOURCE items
                    if isinstance(data, list):
                        # One DESTINATION item
                        if len(selItems) == 1:
                            # If the SOURCE is a list of Keywords, get User confirmation
                            if data[0].nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                                    data1 = unicode(_('Snapshot'), 'utf8')
                                else:
                                    prompt = _('Do you want to add multiple Keywords to %s "%s"?')
                                    data1 = _('Snapshot')
                                # Set up data to go with the prompt
                                promptdata = (data1, self.GetItemText(sel))
                                # Show the user prompt
                                dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                        # Multiple DESTINATION items
                        else:
                            # If the SOURCE is a list of Keywords, get User confirmation
                            if data[0].nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add multiple Keywords to multiple %s?'), 'utf8')
                                    data1 = unicode(_('Snapshots'), 'utf8')
                                else:
                                    prompt = _('Do you want to add multiple Keywords to multiple %s?')
                                    data1 = _('Snapshots')
                                # Set up data to go with the prompt
                                promptdata = (data1,)
                                # Set up data to go with the prompt
                                dlg = Dialogs.QuestionDialog(None, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                    # One SOURCE item
                    else:
                        # One DESTINATION item
                        if len(selItems) == 1:
                            # ... we skip a prompt and indicate the user said Yes
                            result = wx.ID_YES
                            # Prompt gets handled by DragAndDropObjects.DropKeyword().  We DO want user confirmation
                            confirmations = True
                        # Multiple DESTINATION items
                        else:
                            # If the SOURCE is a Keyword, get User confirmation
                            if data.nodetype == 'KeywordNode':
                                # Get user confirmation of the Keyword Add request
                                if 'unicode' in wx.PlatformInfo:
                                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                    prompt = unicode(_('Do you want to add Keyword "%s:%s" to multiple %s?'), 'utf8') 
                                    data1 = unicode(_('Snapshots'), 'utf8')
                                else:
                                    prompt = _('Do you want to add Keyword "%s:%s" to multiple %s?')
                                    data1 = _('Snapshots')
                                # Set up data to go with the prompt
                                promptdata = (data.parent, data.text, data1)
                                # Show the user prompt
                                dlg = Dialogs.QuestionDialog(self.parent, prompt % promptdata)
                                result = dlg.LocalShowModal()
                                dlg.Destroy()
                            # If we don't have a Keyword ...
                            else:
                                # ... we skip a prompt and indicate the user said Yes
                                result = wx.ID_YES
                            # We DO NOT want user confirmation
                            confirmations = False

                    # If the user confirms, or we skip the confirmation programatically ...
                    if result == wx.ID_YES:
                        # For each Clip in the selected items ...
                        for item in selItems:
                            # ... grab the individual item ...
                            sel = item
                            # If data is a list, there are multiple Clip nodes to paste
                            if isinstance(data, list):
                                # Iterate through the Clip nodes
                                for datum in data:
                                    # ... and paste the data
                                    DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'], confirmations=False)
                            # if data is NOT a list, it is a single Clip node to paste
                            else:
                                # ... and paste the data
                                DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'], confirmations=confirmations)
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Snapshot

        elif n == 4:      # Add Quote
            try:
                # Get the Document Selection information from the ControlObject.
                (documentNum, startChar, endChar, text) = self.parent.ControlObject.GetDocumentSelectionInfo()
                # If there's a selection in the text ...
                if text != '':
                    # ... copy it to the clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                    # Add the Quote.
                    self.parent.add_quote(coll_name)
                else:
                    # Create an error message.
                    msg = _('You must make a selection in a document to be able to add a Quote.')
                    errordlg = Dialogs.InfoDialog(None, msg)
                    errordlg.ShowModal()
                    errordlg.Destroy()
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

        elif n == 5:      # Add Clip
            try:
                # Get the Transcript Selection information from the ControlObject.
                (transcriptNum, startTime, endTime, text) = self.parent.ControlObject.GetTranscriptSelectionInfo()
                # If a selection has been made ...
                if text != '':
                    # ... copy that to the Clipboard by faking a Drag event!
                    self.parent.ControlObject.TranscriptWindow.dlg.editor.OnStartDrag(evt, copyToClipboard=True)
                # Add the Clip
                self.parent.add_clip(coll_name)

                # Let's send a message telling others they need to re-order this collection!
                # This is necessary because the Message Server doesn't have the capacity to insert the clip in the right place
                # on remote computers
                if not TransanaConstants.singleUserVersion:
                    # Get the Collection (the Clip's Parent)
                    tempCollection = Collection.Collection(coll_name, coll_parent_num)
                    # Start by getting the Collection's Node Data
                    nodeData = tempCollection.GetNodeData()
                    # Indicate that we're sending a Collection Node
                    msg = "CollectionNode"
                    # Iterate through the nodes
                    for node in nodeData:
                        # ... add the appropriate seperator and then the node name
                        msg += ' >|< %s' % node

                    if DEBUG:
                        print 'Message to send = "OC %s"' % msg

                    # Send the Order Collection Node message
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("OC %s" % msg)

            except:

                print sys.exc_info()[0]
                print sys.exc_info()[1]
                import traceback
                traceback.print_exc(file=sys.stdout)
                
                msg = _('You must make a selection in a transcript to be able to add a Clip.')
                errordlg = Dialogs.InfoDialog(None, msg)
                errordlg.ShowModal()
                errordlg.Destroy()

        elif n == 6:      # Add Multi-transcript Clip
            # Create the appropriate Clip object
            tempClip = self.CreateMultiTranscriptClip(selData.parent, coll_name)
            # Add the Clip, using EditClip because we don't want the overhead of the Drag-and-Drop architecture
            self.parent.edit_clip(tempClip)
            # Get the Collection info
            tempCollection = Collection.Collection(tempClip.collection_num)
            # In this case, the clip has been assigned the WRONG Sort Order.  Let's wipe it out to signal that the
            # ChangeClipOrder call should reposition the clip

            tempClip.lock_record()
            tempClip.sort_order = 0
            tempClip.db_save()
            tempClip.unlock_record()
            
            # Now change the Sort Order, and if it succeeds ...
            if DragAndDropObjects.ChangeClipOrder(self, sel, tempClip, tempCollection):
                # ... let's send a message telling others they need to re-order this collection!
                if not TransanaConstants.singleUserVersion:
                    # Start by getting the Collection's Node Data
                    nodeData = tempCollection.GetNodeData()
                    # Indicate that we're sending a Collection Node
                    msg = "CollectionNode"
                    # Iterate through the nodes
                    for node in nodeData:
                        # ... add the appropriate seperator and then the node name
                        msg += ' >|< %s' % node

                    if DEBUG:
                        print 'Message to send = "OC %s"' % msg

                    # Send the Order Collection Node message
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("OC %s" % msg)

        elif n == 7:      # Add Snapshot
            # Get the Collection (the Clip's Parent)
            coll = Collection.Collection(coll_name, coll_parent_num)
            # Call the Add Snapshot dialog with the current Collection Number
            self.parent.add_snapshot(coll.number)

        elif n == 8:      # Add Snapshot Note
            self.parent.add_note(snapshotNum=selData.recNum)

        elif n == 9:      # Delete
            # Notes Browser Check and Adjustment            
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Snapshot Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # For each Snapshot in the selected items ...
                for item in selItems:
                    self.DeleteQuoteClipSnapshotItem(item)
                    
##                    # ... grab the individual item ...
##                    sel = item
##                    # ... and get the item's data
##                    selData = self.GetPyData(sel)
##
##                    # Load the Selected Snapshot, signalling that we intend to delete it
##                    tmpSnapshot = Snapshot.Snapshot(selData.recNum, suppressEpisodeError = True)
##                    # Remember the Snapshot Number, to use after the Snapshot has been deleted
##                    snapshotNum = tmpSnapshot.number
##                    # Get user confirmation of the Snapshot Delete request
##                    if 'unicode' in wx.PlatformInfo:
##                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                        prompt = unicode(_('Are you sure you want to delete Snapshot "%s" and all related Notes?'), 'utf8')
##                    else:
##                        prompt = _('Are you sure you want to delete Snapshot "%s" and all related Notes?')
##                    dlg = Dialogs.QuestionDialog(self, prompt % self.GetItemText(sel))
##                    result = dlg.LocalShowModal()
##                    dlg.Destroy()
##
##                    # If the user confirms the Delete Request...
##                    if result == wx.ID_YES:
##
##                        # Determine what Keyword Examples exist for the specified Snapshot so that they can be removed from the
##                        # Database Tree if the delete succeeds.  We must do that first, as the Snapshots themselves and the
##                        # ClipKeywords Records will be deleted later!
####                        kwExamples = DBInterface.list_all_keyword_examples_for_a_snapshot(snapshotNum)
##
###                        print "DatabaseTreeTab.OnSnapshotCommand() -- Deleting Snapshot:"
###                        print "  Need to delete Keyword Example Nodes !!"
###                        print
##
##                        # We need to remember the Episode Number after the Snapshot is deleted
##                        tmpEpisodeNum = tmpSnapshot.episode_num
##
##                        try:
##                            # Get the Snapshot's Node List
##                            nodeList = (_("Collections"), ) + tmpSnapshot.GetNodeData()
##                            # Try to delete the Snapshot, initiating a Transaction
##                            delResult = tmpSnapshot.db_delete(1)
##                            # If successful, remove the Clip Node from the Database Tree
##                            if delResult:
###                                # Get a temporary Selection Pointer
###                                tempSel = sel
###                                # Get the full Node Branch by climbing it to one level above the root
###                                nodeList = (self.GetItemText(tempSel),)
###                                while (self.GetItemParent(tempSel) != self.GetRootItem()):
###                                    tempSel = self.GetItemParent(tempSel)
###                                    nodeList = (self.GetItemText(tempSel),) + nodeList
##
##                                # Deleting all these ClipKeyword records needs to remove Keyword Example Nodes in the DBTree.
##                                # That needs to be done here in the User Interface rather than in the Clip Object, as that is
##                                # a user interface issue.  The Clip Record and the Clip Keywords Records get deleted, but
##                                # the user interface does not get cleaned up by deleting the Clip Object.
###                                for (kwg, kw, clipNum, clipID) in kwExamples:
###                                    self.delete_Node((_("Keywords"), kwg, kw, clipID), 'KeywordExampleNode', exampleClipNum = clipNum)
##
##                                # Call the DB Tree's delete_Node method.
##                                self.delete_Node(nodeList, 'SnapshotNode')
##                                # ... and if the Notes Browser is open, ...
##                                if self.parent.ControlObject.NotesBrowserWindow != None:
##                                    # ... we need to CHECK to see if any notes were deleted.
##                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
##
##                                # If the Snapshot was attached to an Episode ...
##                                if tmpEpisodeNum > 0:
##                                    # ... see if the Keyword visualization needs to be updated.
##                                    self.parent.ControlObject.UpdateKeywordVisualization()
##                                    # Even if this computer doesn't need to update the keyword visualization others, might need to.
##                                    if not TransanaConstants.singleUserVersion:
##                                        # We need to update the Episode Keyword Visualization
##                                        if TransanaGlobal.chatWindow != None:
##                                            TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', tmpEpisodeNum, 0))
##                                    
##
##                        # Handle the RecordLocked exception, which arises when records are locked!
##                        except RecordLockedError, e:
##                            # Display the Exception Message, allow "continue" flag to remain true
##                            if 'unicode' in wx.PlatformInfo:
##                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                                prompt = unicode(_('You cannot delete Snapshot "%s".\n%s'), 'utf8')
##                            else:
##                                prompt = _('You cannot delete Snapshot "%s".\n%s')
##                            errordlg = Dialogs.ErrorDialog(None, prompt % (tmpSnapshot.id, e.explanation))
##                            errordlg.ShowModal()
##                            errordlg.Destroy()
##                        # Handle other exceptions
##                        except:
##                            if DEBUG:
##                                import traceback
##                                traceback.print_exc(file=sys.stdout)
##
##                            # Display the Exception Message, allow "continue" flag to remain true
##                            prompt = "%s : %s"
##                            if 'unicode' in wx.PlatformInfo:
##                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
##                                prompt = unicode(prompt, 'utf8')
##                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
##                            errordlg.ShowModal()
##                            errordlg.Destroy()

        elif n == 10:      # Load Snapshot Context
            # Load the Snapshot.
            snapshot = Snapshot.Snapshot(selData.recNum)
            # If the Snapshot HAS a defined context (and didn't have problems loading Episode or Transcript) ...
            if (snapshot.episode_num > 0) and (snapshot.transcript_num > 0):
                # Load the Episode
                episode = Episode.Episode(snapshot.episode_num)
                # Initialize a list for Transcript Numbers
                transcriptNums = []
                # Get a list of possible transcripts
                transcripts = DBInterface.list_transcripts(episode.series_id, episode.id)
                # For each Transcript in the Transcript List ...
                for tr in transcripts:
                    # ... add the transcript Number to the Transcript List
                    transcriptNums.append(tr[0])
                # Start exception handling to catch failures due to orphaned clips
                try:
                    # If the transcript is in the Transcript list ...
                    if snapshot.transcript_num in transcriptNums:
                        # ... load the transcript
                        # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                        transcript = Transcript.Transcript(snapshot.transcript_num, skipText=True)
                    # If any of the transcripts are orphans ...
                    else:
    ####                    # Get the list of possible replacement source transcripts
    ####                    transcriptList = DBInterface.list_transcripts(episode.series_id, episode.id)
    ####                    # If only 1 transcript is in the list ...
    ####                    if len(transcriptList) == 1:
    ####                        # ... use that.
    ####                        transcript = Transcript.Transcript(transcriptList[0][0])
    ####                    # If there are NO transcripts (perhaps because the Episode is gone, perhaps because it has no Transcripts) ...
    ####                    elif len(transcriptList) == 0:
    ####                        # ... raise an exception.  We can't locate a transcript if the Episode is gone or if it has no Transcripts.
    ####                        raise RecordNotFoundError ('Transcript', 0)
    ####                    # If there are multiple transcripts to choose from ...
    ####                    else:
    ####                        # Initialize a list
    ####                        strList = []
    ####                        # Iterate through the list of transcripts ...
    ####                        for (transcriptNum, transcriptID, episodeNum) in transcriptList:
    ####                            # ... and extract the Transcript IDs
    ####                            strList.append(transcriptID)
    ####                        # Create a dialog where the user can choose one Transcript
    ####                        dlg = wx.SingleChoiceDialog(self, _('Transana cannot identify the Transcript where this clip originated.\nPlease select the Transcript that was used to create this Clip.'),
    ####                                                    _('Transana Information'), strList, wx.OK | wx.CANCEL)
    ####                        # Show the dialog.  If the user presses OK ...
    ####                        if dlg.ShowModal() == wx.ID_OK:
    ####                            # ... use the selected transcript
    ####                            transcript = Transcript.Transcript(dlg.GetStringSelection(), episode.number)
    ####                        # If the user presses Cancel (Esc, etc.) ...
    ####                        else:
                                # ... raise an exception
                                raise RecordNotFoundError ('Transcript', snapshot.transcript_num)
                    # As long as a Control Object is defined (which it always will be)
                    if self.parent.ControlObject != None:
                        # Set the active transcript to 0 so the whole interface will be reset
                        self.parent.ControlObject.activeTranscript = 0
                        # Load the source Transcript
                        self.parent.ControlObject.LoadTranscript(episode.series_id, episode.id, transcript.id)
                        # Check to see if the load succeeded before continuing!
                        if self.parent.ControlObject.currentObj != None:
                            # We need to signal that the Visualization needs to be re-drawn.
                            self.parent.ControlObject.ChangeVisualization()
                            # Mark the Clip as the current selection.  (This needs to be done AFTER all transcripts have been opened.)
                            self.parent.ControlObject.SetVideoSelection(snapshot.episode_start, snapshot.episode_start + snapshot.episode_duration)
                            # We need the screen to update here, before the next step.
                            wx.Yield()
                            # Now let's go through each Transcript Window ...
                            for trWin in self.parent.ControlObject.TranscriptWindow:
                                # ... move the cursor to the TRANSCRIPT's Start Time (not the Clip's)
                                trWin.dlg.editor.scroll_to_time(snapshot.episode_start + 500)
                                # .. and select to the TRANSCRIPT's End Time (not the Clip's)
                                trWin.dlg.editor.select_find(str(snapshot.episode_start + snapshot.episode_duration))
                                # update the selection text
                                wx.CallLater(50, trWin.dlg.editor.ShowCurrentSelection)
                except:
                    (exctype, excvalue, traceback) = sys.exc_info()

                    if DEBUG:
                        print exctype, excvalue
                        import traceback
                        traceback.print_exc(file=sys.stdout)
                    
                    if len(transcripts) == 1:
                        msg = _('The Transcript this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
                    else:
                        msg = _('One of the Transcripts this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
                    self.parent.ControlObject.ClearAllWindows()
                    dlg = Dialogs.ErrorDialog(self, msg)
                    result = dlg.ShowModal()
                    dlg.Destroy()
            else:
                msg = unicode(_('Snapshot "%s" does not have a defined context.\nLibrary ID, Episode ID, and Transcript ID must all be selected.'), 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg % snapshot.id)
                result = dlg.ShowModal()
                dlg.Destroy()
                # Load the Snapshot Properties form
                self.parent.edit_snapshot(snapshot)

        elif n == 11:  # Insert Snapshot Hyperlink
            # If the current transcript is in Read Only mode ...
            if self.parent.ControlObject.ActiveTranscriptReadOnly():
                # ... inform the user
                msg = _("The current document is not editable.  The requested snapshot cannot be inserted into the document.")
                msg += '\n\n' + _("To insert the snapshot into the document, press the Edit Mode button on the Document Toolbar to make the document editable.")
                dlg = Dialogs.InfoDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
            else:
                self.parent.ControlObject.TranscriptInsertHyperlink('Snapshot', selData.recNum)

        elif n == 12:      # Snapshot Properties
            # Load the Snapshot.
            snapshot = Snapshot.Snapshot(selData.recNum)
            # Load the Snapshot Properties form
            self.parent.edit_snapshot(snapshot)

    def DeleteQuoteClipSnapshotItem(self, item):
        """ Handle the shared deletion of Quote, Clip, and Snapshot items """
        # ... grab the individual item ...
        sel = item
        # ... and get the item's data
        selData = self.GetPyData(sel)

        if self.GetPyData(item).nodetype == 'QuoteNode':
            # Load the Selected Quote.  We DO NOT need the Quote Text here
            tmpObj = Quote.Quote(num=selData.recNum, skipText=True)
            # Get the Object Parent number (so it will persist after the Object is deleted)
            tmpObjParentNum = tmpObj.source_document_num
            tmpObjType = Quote.Quote
            tmpObjTypeText = 'Quote'
            # Get user confirmation of the Quote Delete request
            prompt = unicode(_('Are you sure you want to delete Quote "%s" and all related Notes?'), 'utf8')
            delPrompt = unicode(_('You cannot delete Quote "%s".\n%s'), 'utf8')
        elif self.GetPyData(item).nodetype == 'ClipNode':
            # Load the Selected Clip.  We DO need the Clip Text here, otherwise we can't delete because the
            # Clip Transcript doesn't get locked.
            tmpObj = Clip.Clip(selData.recNum)
            # Get the Object Parent number (so it will persist after the Object is deleted)
            tmpObjParentNum = tmpObj.episode_num
            tmpObjType = Clip.Clip
            tmpObjTypeText = 'Clip'
            # Get user confirmation of the Clip Delete request
            prompt = unicode(_('Are you sure you want to delete Clip "%s" and all related Notes?'), 'utf8')
            delPrompt = unicode(_('You cannot delete Clip "%s".\n%s'), 'utf8')
        elif self.GetPyData(item).nodetype == 'SnapshotNode':
            # Load the Selected Snapshot.
            tmpObj = Snapshot.Snapshot(selData.recNum)
            # There is no Parent Object number
            tmpObjParentNum = tmpObj.episode_num
            tmpObjType = Snapshot.Snapshot
            tmpObjTypeText = 'Snapshot'
            # Get user confirmation of the Snapshot Delete request
            prompt = unicode(_('Are you sure you want to delete Snapshot "%s" and all related Notes?'), 'utf8')
            delPrompt = unicode(_('You cannot delete Snapshot "%s".\n%s'), 'utf8')
                
        dlg = Dialogs.QuestionDialog(self, prompt % self.GetItemText(sel))
        result = dlg.LocalShowModal()
        dlg.Destroy()

        # If the user confirms the Delete Request...
        if result == wx.ID_YES:
            # If THIS Object is currently loaded, we need to remove it from the TranscriptWindow
            self.parent.ControlObject.CloseOpenTranscriptWindowObject(tmpObjType, tmpObj.number)

            if isinstance(tmpObj, Clip.Clip):
                # Determine what Keyword Examples exist for the specified Clip so that they can be removed from the
                # Database Tree if the delete succeeds.  We must do that first, as the Clips themselves and the
                # ClipKeywords Records will be deleted later!
                kwExamples = DBInterface.list_all_keyword_examples_for_a_clip(selData.recNum)

            # Start exception handling
            try:
                # Get the Object number (so it will persist after the Object is deleted)
                tmpObjNum = tmpObj.number
                # If we have a Quote object ...
                if isinstance(tmpObj, Quote.Quote):
                    # Get the Source Document's Number
                    tmpObjSourceNum = tmpObj.source_document_num
                # Try to delete the Object, initiating a Transaction
                delResult = tmpObj.db_delete(1)
                # If successful, remove the Object Node from the Database Tree
                if delResult:
                    if isinstance(tmpObj, Quote.Quote):
                        # Remove the Object's Position Data from the Source Document, if it's open
                        self.parent.ControlObject.RemoveQuoteFromOpenDocument(tmpObjNum, tmpObjParentNum)

                        # Even if this computer doesn't need to update the Source Document, others might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to pass the type of the deleted Object's record number and the deleted Object's
                            # SOURCE Document number.
                            if DEBUG:
                                print 'Message to send = "DQPOD %s %s"' % (tmpObjNum, tmpObjSourceNum)
                                
                            if (TransanaGlobal.chatWindow != None) and (tmpObjSourceNum > 0):
                                TransanaGlobal.chatWindow.SendMessage("DQPOD %s %s" % (tmpObjNum, tmpObjSourceNum))

                    # Get a temporary Selection Pointer
                    tempSel = sel
                    # Get the full Node Branch by climbing it to one level above the root
                    nodeList = (self.GetItemText(tempSel),)
                    while (self.GetItemParent(tempSel) != self.GetRootItem()):
                        tempSel = self.GetItemParent(tempSel)
                        nodeList = (self.GetItemText(tempSel),) + nodeList

                    if isinstance(tmpObj, Clip.Clip):
                        # Deleting all these ClipKeyword records needs to remove Keyword Example Nodes in the DBTree.
                        # That needs to be done here in the User Interface rather than in the Clip Object, as that is
                        # a user interface issue.  The Clip Record and the Clip Keywords Records get deleted, but
                        # the user interface does not get cleaned up by deleting the Clip Object.
                        for (kwg, kw, clipNum, clipID) in kwExamples:
                            self.delete_Node((_("Keywords"), kwg, kw, clipID), 'KeywordExampleNode', exampleClipNum = clipNum)

                    # Call the DB Tree's delete_Node method.
                    self.delete_Node(nodeList, self.GetPyData(item).nodetype)
                    # ... and if the Notes Browser is open, ...
                    if self.parent.ControlObject.NotesBrowserWindow != None:
                        # ... we need to CHECK to see if any notes were deleted.
                        self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('C')
                        
                    # If the current main object is a Document and it's the Document that contains the
                    # deleted Object, we need to update the Document Keyword Visualization!
                    if (isinstance(self.parent.ControlObject.currentObj, Document.Document)) and \
                       (tmpObjParentNum == self.parent.ControlObject.currentObj.number):
                        self.parent.ControlObject.UpdateKeywordVisualization()

                    # If the current main object is an Episode and it's the Episode that contains the
                    # deleted Object, we need to update the Episdoe Keyword Visualization!
                    if (isinstance(self.parent.ControlObject.currentObj, Episode.Episode)) and \
                       (tmpObjParentNum == self.parent.ControlObject.currentObj.number):
                        self.parent.ControlObject.UpdateKeywordVisualization()

                    if isinstance(self.parent.ControlObject.currentObj, Snapshot.Snapshot) and \
                        (tmpObjParentNum > 0):
                        # ... see if the Keyword visualization needs to be updated.
                        self.parent.ControlObject.UpdateKeywordVisualization()
                        # Even if this computer doesn't need to update the keyword visualization others, might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to update the Episode Keyword Visualization
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', tmpObjParentNum, 0))

                    if isinstance(self.parent.ControlObject.currentObj, Episode.Episode) or \
                       isinstance(self.parent.ControlObject.currentObj, Document.Document):
                        # Even if this computer doesn't need to update the keyword visualization others, might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to pass the type of the current object, the deleted Object's record number, and
                            # the deleted Object's Parent number.
                            if DEBUG:
                                print 'Message to send = "UKV %s %s %s"' % (tmpObjTypeText, tmpObjNum, tmpObjParentNum)
                                
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % (tmpObjTypeText, tmpObjNum, tmpObjParentNum))

            # Handle the RecordLocked exception, which arises when records are locked!
            except RecordLockedError, e:
                # Display the Exception Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, delPrompt % (tmpObj.id, e.explanation))
                errordlg.ShowModal()
                errordlg.Destroy()
            # Handle other exceptions
            except:
                if DEBUG:
                    import traceback
                    traceback.print_exc(file=sys.stdout)

                # Display the Exception Message, allow "continue" flag to remain true
                prompt = "%s : %s"
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(prompt, 'utf8')
                errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()

    def OnNoteCommand(self, evt):
        """Handle selections for the Note menu."""
        n = evt.GetId() - self.cmd_id_start["NoteNode"]
        
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            note_name = self.GetItemText(sel)
        
        if n == 0:      # Cut
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Note Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
                # If there's only one selected item ...
                if len(selItems) == 1:
                    self.cutCopyInfo['sourceItem'] = sel
                else:
                    self.cutCopyInfo['sourceItem'] = selItems
                self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Note

        elif n == 3:    # Open in Notes Browser
            # If the NotesBrowser is NOT currently open ...
            if self.parent.ControlObject.NotesBrowserWindow == None:
                # Instantiate a Notes Browser window
                notesBrowser = NotesBrowser.NotesBrowser(self.parent.ControlObject.MenuWindow, -1, _("Notes Browser"))
                # Register the Control Object with the Notes Browser
                notesBrowser.Register(self.parent.ControlObject)
                # Display the Notes Browser
                notesBrowser.Show()
                # Open the appropriate note
                notesBrowser.OpenNote(selData.recNum)
            # If the Notes Browser IS currently open ...
            else:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Open the appropriate note
                self.parent.ControlObject.NotesBrowserWindow.OpenNote(selData.recNum)

        elif n == 4:    # Delete
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Note Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # For each Note in the selected items ...
                for item in selItems:
                    # ... grab the individual item ...
                    sel = item
                    # ... and get the item's data
                    selData = self.GetPyData(sel)

                    # Load the Selected Note
                    note = Note.Note(selData.recNum)
                    # Get user confirmation of the Note Delete request
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Are you sure you want to delete Note "%s"?'), 'utf8')
                    else:
                        prompt = _('Are you sure you want to delete Note "%s"?')
                    dlg = Dialogs.QuestionDialog(self, prompt % (self.GetItemText(sel)))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                    # If the user confirms the Delete Request...
                    if result == wx.ID_YES:
                        try:
                            # Try to delete the Note.  There is no need to initiate a Transaction for deleting a Note.
                            delResult = note.db_delete(0)
                            # If successful, remove the Note Node from the Database Tree
                            if delResult:
                                # Get the full Node Branch by climbing it to one level above the root
                                nodeList = (self.GetItemText(sel),)
                                while (self.GetItemParent(sel) != self.GetRootItem()):
                                    sel = self.GetItemParent(sel)
                                    nodeList = (self.GetItemText(sel),) + nodeList
                                sel = item
                                # Call the DB Tree's delete_Node method.
                                # To climb the DBTree properly, we need to provide the note's PARENT's NodeType along with the NodeNode indication
                                noteParentNodeType = self.GetPyData(self.GetItemParent(sel)).nodetype
                                if noteParentNodeType == 'LibraryNode':
                                    noteNodeType = 'LibraryNoteNode'
                                    nodeType = 'Library'
                                elif noteParentNodeType == 'EpisodeNode':
                                    noteNodeType = 'EpisodeNoteNode'
                                    nodeType = 'Episode'
                                elif noteParentNodeType == 'TranscriptNode':
                                    noteNodeType = 'TranscriptNoteNode'
                                    nodeType = 'Transcript'
                                elif noteParentNodeType == 'CollectionNode':
                                    noteNodeType = 'CollectionNoteNode'
                                    nodeType = 'Collection'
                                elif noteParentNodeType == 'QuoteNode':
                                    noteNodeType = 'QuoteNoteNode'
                                    nodeType = 'Quote'
                                elif noteParentNodeType == 'ClipNode':
                                    noteNodeType = 'ClipNoteNode'
                                    nodeType = 'Clip'
                                elif noteParentNodeType == 'SnapshotNode':
                                    noteNodeType = 'SnapshotNoteNode'
                                    nodeType = 'Snapshot'
                                elif noteParentNodeType == 'DocumentNode':
                                    noteNodeType = 'DocumentNoteNode'
                                    nodeType = 'Document'
                                # Delete the tree node
                                self.delete_Node(nodeList, noteNodeType)
                                # if the Notes Browser is open ...
                                if self.parent.ControlObject.NotesBrowserWindow != None:
                                    # ... remove the Note from the Notes Browser
                                    self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('D', (nodeType, nodeList))
                        # Handle the RecordLocked exception, which arises when records are locked!
                        except RecordLockedError, e:
                            # Display the Exception Message, allow "continue" flag to remain true
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(_('You cannot delete Note "%s".\n%s'), 'utf8')
                            else:
                                prompt = _('You cannot delete Note "%s".\n%s')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (note.id, e.explanation))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                        # Handle other exceptions
                        except:
                            # Display the Exception Message, allow "continue" flag to remain true
                            prompt = "%s : %s"
                            if 'unicode' in wx.PlatformInfo:
                                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                                prompt = unicode(prompt, 'utf8')
                            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                            errordlg.ShowModal()
                            errordlg.Destroy()

        elif n == 5:  # Insert Note Hyperlink
            # If the current transcript is in Read Only mode ...
            if self.parent.ControlObject.ActiveTranscriptReadOnly():
                # ... inform the user
                msg = _("The current document is not editable.  The requested hyperlink cannot be inserted into the document.")
                msg += '\n\n' + _("To insert the hyperlink into the document, press the Edit Mode button on the Document Toolbar to make the document editable.")
                dlg = Dialogs.InfoDialog(self, msg)
                dlg.ShowModal()
                dlg.Destroy()
            else:
                # Insert the Hyperlink
                self.parent.ControlObject.TranscriptInsertHyperlink("Note", selData.recNum)

        elif n == 6:    # Note Properties
            if (self.parent.ControlObject.NotesBrowserWindow != None) and TransanaConstants.singleUserVersion:
                # ... make it visible, on top of other windows
                self.parent.ControlObject.NotesBrowserWindow.Raise()
                # If the window has been minimized ...
                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                    # ... then restore it to its proper size!
                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                # Get user confirmation of the Note Delete request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?'), 'utf8')
                else:
                    prompt = _('This option is not available while the Notes Browser is open.\nClose the Notes Browser and continue?')
                dlg = Dialogs.QuestionDialog(self, prompt)
                result = dlg.LocalShowModal()
                dlg.Destroy()
                if result == wx.ID_YES:
                    self.parent.ControlObject.NotesBrowserWindow.Close()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                # Load the Note
                note = Note.Note(selData.recNum)
                # Edit the Note's Properties
                self.parent.edit_note(note)

        else:
            raise MenuIDError

    def updateKWGroupsData(self):
        """ Refresh internal data about keyword groups """
        # Since we've just inserted a new Keyword Group, we need to rebuild the self.kwgroups data structure.
        # This data structure is used to ensure that empty keyword groups still show up in the Keyword Properties dialog.
        # Initialize keyword groups to an empty list
        self.kwgroups = []
        # The "Keywords" node itself is always item 0 in the node list
        kwg_root = self.select_Node((_("Keywords"),), 'KeywordRootNode')
        self.kwgroups.append(kwg_root)
        (child, cookieVal) = self.GetFirstChild(kwg_root)
        while child.IsOk():
            self.kwgroups.append(child)
            (child, cookieVal) = self.GetNextChild(kwg_root, cookieVal)

    def OnKwRootCommand(self, evt):
        """Handle selections for the root Keyword group menu."""
        n = evt.GetId() - self.cmd_id_start["KeywordRootNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
        
        if n == 0:      # Add KW group
            # Get list of keyword group names in the tree.
            # The database may not contain everything in the tree due to the
            # database structure (keyword groups are a field of keyword
            # records, and dont have a record in itself).
            kwg_names = []
            for kw in self.kwgroups[1:]:        # skip 'Keywords' node @ 0
                kwg_names.append(self.GetItemText(kw))
     
            kwgDlg = Dialogs.add_kw_group_ui(self, kwg_names)
            result = kwgDlg.ShowModal()
            if result == wx.ID_OK:
                kwg = kwgDlg.kwGroup.GetValue()
            else:
                kwg = None
            kwgDlg.Destroy()
            
            if kwg:
                nodeData = (_('Keywords'), kwg)
                # Add the new Keyword Group to the data tree
                self.add_Node('KeywordGroupNode', nodeData, 0, 0)

                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    if DEBUG:
                        print 'Message to send = "AKG %s"' % nodeData[-1]
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("AKG %s" % nodeData[-1])

                self.updateKWGroupsData()

        elif n == 1:    # KW Management
            # Call up the Keyword Management dialog
            KWManager.KWManager(self)
            # Refresh the Keywords Node to show the changes that were made
            self.refresh_kwgroups_node()
            # We need to update the Keyword Visualization! (Well, maybe not, but I don't know how to tell!)
            self.parent.ControlObject.UpdateKeywordVisualization()

        elif n == 2:    # Keyword Summary Report
            if self.ItemHasChildren(sel):
                KeywordSummaryReport.KeywordSummaryReport(controlObject = self.parent.ControlObject)
            else:
                # Display the Error Message
                dlg = Dialogs.ErrorDialog(None, _('The requested Keyword Summary Report contains no data to display.'))
                dlg.ShowModal()
                dlg.Destroy()

        else:
            raise MenuIDError
  
    def OnKwGroupCommand(self, evt):
        """Handle selections for the root Keyword group menu."""
        n = evt.GetId() - self.cmd_id_start["KeywordGroupNode"]
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... and set some variables
            kwg_name = self.GetItemText(sel)
        
        if n == 0:      # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # We can't MOVE to multiple locations!
            if len(selItems) > 1:
                self.cutCopyInfo['action'] = 'Copy'
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                # If data is a list, there are multiple Clip nodes to paste
                if isinstance(data, list):
                    # Iterate through the Clip nodes
                    for datum in data:
                        # Iterate through the selected nodes too
                        for sel in selItems:
                        # ... and paste the data
                            DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                    # if data is NOT a list, it is a single Clip node to paste
                else:
                    # Iterate through the selected nodes too
                    for sel in selItems:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 1:    # Add Keyword
            self.parent.add_keyword(kwg_name)

        elif n == 2:    # Delete this keyword group
            # For each Keyword Group in the selected items ...
            for item in selItems:
                # ... grab the individual item ...
                sel = item
                # ... and get the item's data
                selData = self.GetPyData(sel)
                kwg_name = self.GetItemText(sel)

                # We need to update all open Snapshot Windows based on the change in this Keyword
                # If we found the keyword we're supposed to delete ...
                if self.parent.ControlObject.UpdateSnapshotWindows('Detect', evt, kwg_name):
                    # ... present an error message to the user
                    msg = _("A Keyword from Keyword Group %s is contained in a Snapshot you are currently editing.\nYou cannot delete it at this time.")
                    msg = unicode(msg, 'utf8')
                    dlg = Dialogs.ErrorDialog(self, msg % (kwg_name))
                    dlg.ShowModal()
                    dlg.Destroy()
                    # Signal that we do NOT want to delete the Keyword!
                    result = wx.ID_NO
                else:
                    msg = _('Are you sure you want to delete Keyword Group "%s", all of its keywords, and all instances of those keywords from Quotes, Clips, and Snapshots?')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                    dlg = Dialogs.QuestionDialog(self, msg % kwg_name)
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                if result == wx.ID_YES:
                    try:
                        # Delete the Keyword group
                        DBInterface.delete_keyword_group(kwg_name)
                        # Remove the Keyword Group from the tree
                        # Get the full Node Branch by climbing it to one level above the root
                        nodeList = (self.GetItemText(sel),)
                        while (self.GetItemParent(sel) != self.GetRootItem()):
                            sel = self.GetItemParent(sel)
                            nodeList = (self.GetItemText(sel),) + nodeList
                        # Call the DB Tree's delete_Node method.
                        self.delete_Node(nodeList, 'KeywordGroupNode')
                        # We maintain a list of keyword groups so that empty ones don't get lost.  We need to remove the deleted keyword group from 
                        # this list as well
                        # NOTE: "self.kwgroups.remove(sel)" doesn't work if we've been messing with KWG capitalization ("Test : 1" and "test : 2" appear in the same KWG.)
                        self.updateKWGroupsData()

                        # We need to update all open Snapshot Windows based on the change in this Keyword
                        self.parent.ControlObject.UpdateSnapshotWindows('Update', evt, kwg_name)
                        # We need to update the Keyword Visualization!  (Well, maybe not, but I don't know how to tell!)
                        self.parent.ControlObject.UpdateKeywordVisualization()
                        # Even if this computer doesn't need to update the keyword visualization others, might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to update the Keyword Visualization no matter what here, when deleting a keyword group
                            if DEBUG:
                                print 'Message to send = "UKV %s %s %s"' % ('None', 0, 0)
                                
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))

                    # Handle the RecordLocked exception, which arises when records are locked!
                    except RecordLockedError, e:
                        # Display the Exception Message, allow "continue" flag to remain true
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot delete Keyword Group "%s".\n%s'), 'utf8')
                        else:
                            prompt = _('You cannot delete Keyword Group "%s".\n%s')
                        errordlg = Dialogs.ErrorDialog(None, prompt % (kwg_name, e.explanation))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                    # Handle GeneralError exceptions
                    except GeneralError, e:
                        # Display the Exception Message, allow "continue" flag to remain true
                        prompt = "%s"
                        data = e.explanation
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(prompt, 'utf8')
                        errordlg = Dialogs.ErrorDialog(None, prompt % data)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                    # Handle other exceptions
                    except:

                        if DEBUG:
                            print "Exception %s: %s" % (sys.exc_info()[0], sys.exc_info()[1])
                            import traceback
                            traceback.print_exc(file=sys.stdout)
                            
                        # Display the Exception Message, allow "continue" flag to remain true
                        prompt = "%s"
                        data = sys.exc_info()[1]
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(prompt, 'utf8')
                        errordlg = Dialogs.ErrorDialog(None, prompt % data)
                        errordlg.ShowModal()
                        errordlg.Destroy()

        elif n == 3:    # Keyword Summary Report
            KeywordSummaryReport.KeywordSummaryReport(keywordGroupName = kwg_name, controlObject = self.parent.ControlObject)

        else:
            raise MenuIDError
 
    def OnKwCommand(self, evt):
        """Handle selections for the Keyword menu."""
        n = evt.GetId() - self.cmd_id_start["KeywordNode"]
        # If we're in the Standard version, we need to adjust the menu numbers
        # for Add Quote (4) and Create Multi-transcript Quick Clip (6)
        if not TransanaConstants.proVersion and (n >= 4):
            n += 1
        if not TransanaConstants.proVersion and (n >= 6):
            n += 1
        
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... and set some variables
            kw_group = self.GetItemText(self.GetItemParent(sel))
            kw_name = self.GetItemText(sel)
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:    # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:    # Delete this keyword
            # For each Keyword in the selected items ...
            for item in selItems:
                # ... grab the individual item ...
                sel = item
                # ... and get the item's data
                selData = self.GetPyData(sel)
                kw_group = self.GetItemText(self.GetItemParent(sel))
                kw_name = self.GetItemText(sel)

                # If we found the keyword we're supposed to delete ...
                if self.parent.ControlObject.UpdateSnapshotWindows('Detect', evt, kw_group, kw_name):
                    # ... present an error message to the user
                    msg = _("Keyword %s : %s is contained in a Snapshot you are currently editing.\nYou cannot delete it at this time.")
                    msg = unicode(msg, 'utf8')
                    dlg = Dialogs.ErrorDialog(self, msg % (kw_group, kw_name))
                    dlg.ShowModal()
                    dlg.Destroy()
                    # Signal that we do NOT want to delete the Keyword!
                    result = wx.ID_NO
                else:
                    msg = _('Are you sure you want to delete Keyword "%s : %s" and all instances of it from Quotes, Clips, and Snapshots?')
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        msg = unicode(msg, 'utf8')
                    dlg = Dialogs.QuestionDialog(self, msg % (kw_group, kw_name))
                    result = dlg.LocalShowModal()
                    dlg.Destroy()
                if result == wx.ID_YES:
                    try:
                        # Delete the Keyword
                        DBInterface.delete_keyword(kw_group, kw_name)
                        # Get the full Node Branch by climbing it to one level above the root
                        nodeList = (self.GetItemText(sel),)
                        while (self.GetItemParent(sel) != self.GetRootItem()):
                            sel = self.GetItemParent(sel)
                            nodeList = (self.GetItemText(sel),) + nodeList
                        # Call the DB Tree's delete_Node method.
                        self.delete_Node(nodeList, 'KeywordNode')
                        # We need to update the Keyword Visualization! (Well, maybe not, but I don't know how to tell!)
                        self.parent.ControlObject.UpdateKeywordVisualization()

                        # We need to update all open Snapshot Windows based on the change in this Keyword
                        self.parent.ControlObject.UpdateSnapshotWindows('Update', evt, kw_group, kw_name)
                                
                        # Even if this computer doesn't need to update the keyword visualization others, might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to update the Keyword Visualization no matter what here, when deleting a keyword
                            if DEBUG:
                                print 'Message to send = "UKV %s %s %s"' % ('None', 0, 0)
                                
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))
                    # Handle exceptions
                    except:
                        if DEBUG:
                            print "Exception %s: %s" % (sys.exc_info()[0], sys.exc_info()[1])
                            import traceback
                            traceback.print_exc(file=sys.stdout)
                            
                        # Display the Exception Message, allow "continue" flag to remain true
                        errmsg = sys.exc_info()[1]
                        # If this is a TransanaExceptions.GeneralError, we need the "explanation" parameter!  This should already be Unicode!
                        if hasattr(errmsg, 'explanation'):
                            errmsg = errmsg.explanation
                        elif hasattr(errmsg, 'args'):
                            errmsg = errmsg.args
                        errordlg = Dialogs.ErrorDialog(None, errmsg)
                        errordlg.ShowModal()
                        errordlg.Destroy()

        elif n == 4:    # Create Quick Quote
            try:
                # Get the Document Selection information from the ControlObject.
                (documentNum, startChar, endChar, text) = self.parent.ControlObject.GetDocumentSelectionInfo()
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
            if self.parent.ControlObject.GetCurrentItemType() == 'Document':
                # ... we can just use the ControlObject's Transcript Window's CurrentObj's object number
                sourceDocumentNum = self.parent.ControlObject.TranscriptWindow.GetCurrentObject().number
            # If our source is a Quote ...
            elif self.parent.ControlObject.GetCurrentItemType() == 'Quote':
                # ... we need the ControlObject's currentObj's originating document number
                sourceDocumentNum = self.parent.ControlObject.TranscriptWindow.GetCurrentObject().source_document_num
                startChar += self.parent.ControlObject.TranscriptWindow.GetCurrentObject().start_char
                endChar += self.parent.ControlObject.TranscriptWindow.GetCurrentObject().start_char

            # Initialize the Keyword List
            kwList = []
            # For each Keyword in the selected items ...
            for item in selItems:
                # Get the item's data
                selData = self.GetPyData(item)
                kw_group = self.GetItemText(self.GetItemParent(item))
                kw_name = self.GetItemText(item)
                # Add the keyword to the Keyword List
                kwList.append((kw_group, kw_name))

            # We now have enough information to populate a QuoteDragDropData object to pass to the Quote Creation method.
            quoteData = DragAndDropObjects.QuoteDragDropData(documentNum, sourceDocumentNum, startChar, endChar, text)
            # Pass the accumulated data to the CreateQuickQuote method, which is in the DragAndDropObjects module
            # because drag and drop is an alternate way to create a Quick Quote.
            DragAndDropObjects.CreateQuickQuote(quoteData, kwList[0][0], kwList[0][1], self, extraKeywords=kwList[1:])

        elif n == 5:    # Create Quick Clip
            # Get the Transcript Selection information from the ControlObject, since we can't communicate with the
            # TranscriptEditor directly.
            (transcriptNum, startTime, endTime, text) = self.parent.ControlObject.GetTranscriptSelectionInfo()
            # Initialize the Episode Number to 0
            episodeNum = 0
            # If our source is an Episode ...
            if isinstance(self.parent.ControlObject.currentObj, Episode.Episode):
                # ... we can just use the ControlObject's currentObj's object number
                episodeNum = self.parent.ControlObject.currentObj.number
                # If we are at the end of a transcript and there are no later time codes, Stop Time will be -1.
                # This is, of course, incorrect, and we must replace it with the Episode Length.
                if endTime <= 0:
                    endTime = self.parent.ControlObject.currentObj.tape_length
                videoCheckboxData = self.parent.ControlObject.GetVideoCheckboxDataForClips(startTime)
            # If our source is a Clip ...
            elif isinstance(self.parent.ControlObject.currentObj, Clip.Clip):
                # ... we need the ControlObject's currentObj's originating episode number
                episodeNum = self.parent.ControlObject.currentObj.episode_num
                # Sometimes with a clip, we get a startTime of 0 from the TranscriptSelectionInfo() method.
                # This is, of course, incorrect, and we must replace it with the Clip Start Time.
                if startTime == 0:
                    startTime = self.parent.ControlObject.currentObj.clip_start
                # Sometimes with a clip, we get an endTime of 0 from the TranscriptSelectionInfo() method.
                # This is, of course, incorrect, and we must replace it with the Clip Stop Time.
                if endTime <= 0:
                    endTime = self.parent.ControlObject.currentObj.clip_stop

                # We need to determine what EPISODE media files are used in the CLIP we're quick-clipping from.
                # Get temporary Episode and Clip data
                tmpEpisode = Episode.Episode(episodeNum)
                tmpClip = self.parent.ControlObject.currentObj
                # First, let's find out if the EPISODE's main file is included in the Clip.
                if tmpEpisode.media_filename == tmpClip.media_filename:
                    # If so, its video would be checked, and the CLIP's audio value tells us if its audio would be checked.
                    videoCheckboxData = [(True, tmpClip.audio)]
                # If not ...
                else:
                    # ... then neither box would be checked.
                    videoCheckboxData = [(False, False)]
                # Now let's iterate through the EPISODE's additional media files
                for addEpVidData in tmpEpisode.additional_media_files:
                    # If this Episode Additional Media File is the Clip's Main Media File ...
                    if (addEpVidData['filename'] == tmpClip.media_filename):
                        # ... then the Video Checkbox would have been checked, and the object tells us if audio was too.
                        videoCheckboxData.append((True, addEpVidData['audio']))
                    # If the Clip's main media file was already established, we need to check the Episode additional files
                    # against the Clip Additional media files.
                    else:
                        # Assume the video will not be found ...
                        vidFound = False
                        # ... and that the audio should NOT be included.
                        audIncluded = False
                        # Iterate throught the CLIP additional media files ...
                        for addClipVidData in tmpClip.additional_media_files:
                            # If the Clip Media File matches the Episode Media File ...
                            if addClipVidData['filename'] == addEpVidData['filename']:
                                # ... then we've found the Media File!
                                vidFound = True
                                # Note if audio should be included, getting data from the Clip Object
                                audIncluded = addClipVidData['audio']
                                # Stop looking!
                                break
                        # Add info to videoCheckboxData to indicate whether the Episode's media file and audio were checked when
                        # the source clip was created.
                        videoCheckboxData.append((vidFound, audIncluded))

            # Initialize the Keyword List
            kwList = []
            # For each Keyword in the selected items ...
            for item in selItems:
                # Get the item's data
                selData = self.GetPyData(item)
                kw_group = self.GetItemText(self.GetItemParent(item))
                kw_name = self.GetItemText(item)
                # Add the keyword to the Keyword List
                kwList.append((kw_group, kw_name))
            # We now have enough information to populate a ClipDragDropData object to pass to the Clip Creation method.
            clipData = DragAndDropObjects.ClipDragDropData(transcriptNum, episodeNum, startTime, endTime, text, videoCheckboxData=videoCheckboxData)
            # Pass the accumulated data to the CreateQuickClip method, which is in the DragAndDropObjects module
            # because drag and drop is an alternate way to create a Quick Clip.
            DragAndDropObjects.CreateQuickClip(clipData, kwList[0][0], kwList[0][1], self, extraKeywords=kwList[1:])
            
        elif n == 6:    # Create Multi-transcript Quick Clip
            # If our source is an Episode ...
            if isinstance(self.parent.ControlObject.currentObj, Episode.Episode):
                # Get the Quick Clip Collection information
                (coll_num, coll_name, coll_created) = DBInterface.locate_quick_quotes_and_clips_collection()
                # Create the appropriate Clip object
                tempClip = self.CreateMultiTranscriptClip(coll_num, coll_name)
                # If the Clip Object is created ...
                if tempClip != None:
                    # Initialize the Episode Number to 0
                    episodeNum = tempClip.episode_num
                    # Initialize Transcript Number to 0, which signals multi-transcript Quick Clip
                    transcriptNum = 0
                    # Initialize the Keyword List
                    kwList = []
                    # For each Keyword in the selected items ...
                    for item in selItems:
                        # Get the item's data
                        selData = self.GetPyData(item)
                        kw_group = self.GetItemText(self.GetItemParent(item))
                        kw_name = self.GetItemText(item)
                        # Add the keyword to the Keyword List
                        kwList.append((kw_group, kw_name))
                    # We now have enough information to populate a ClipDragDropData object to pass to the Clip Creation method.
                    clipData = DragAndDropObjects.ClipDragDropData(transcriptNum, episodeNum, tempClip.clip_start, tempClip.clip_stop, tempClip, videoCheckboxData=self.parent.ControlObject.GetVideoCheckboxDataForClips(tempClip.clip_start))
                    # Pass the accumulated data to the CreateQuickClip method, which is in the DragAndDropObjects module
                    # because drag and drop is an alternate way to create a Quick Clip.
                    DragAndDropObjects.CreateQuickClip(clipData, kwList[0][0], kwList[0][1], self, extraKeywords=kwList[1:])

        elif n == 7:    # Quick Search
            # request a Search, passing in the keyword group : keyword information
            search = ProcessSearch.ProcessSearch(self, self.searchCount, kwg=kw_group, kw=kw_name)
            # Get the new searchCount Value (which may or may not be changed)
            self.searchCount = search.GetSearchCount()
            self.UnselectAll()
            self.SelectItem(sel)
            self.EnsureVisible(sel)
            
        elif n == 8:    # Keyword Properties
            # We need to update all open Snapshot Windows based on the change in this Keyword
            # If we found the keyword we're supposed to delete ...
            if self.parent.ControlObject.UpdateSnapshotWindows('Detect', evt, kw_group, kw_name):
                # ... present an error message to the user
                msg = _('Keyword "%s : %s" is contained in a Snapshot you are currently editing.\nYou cannot edit it at this time.')
                msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg % (kw_group, kw_name))
                dlg.ShowModal()
                dlg.Destroy()
            else:
                kw = Keyword.Keyword(kw_group, kw_name)

                self.parent.edit_keyword(kw)

                # We need to update all open Snapshot Windows based on the change in this Keyword
                self.parent.ControlObject.UpdateSnapshotWindows('Update', evt, kw_group, kw_name)
                # We just need to update the Keyword Visualization.  There's no way to tell if the changed
                # keyword appears or not from here.
                self.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # When a collection gets deleted, clips from anywhere could go with it.  It's safest
                    # to update the Keyword Visualization no matter what.
                    if DEBUG:
                        print 'Message to send = "UKV %s %s %s"' % ('None', 0, 0)
                        
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))

        else:
            raise MenuIDError

    def OnKwExampleCommand(self, evt):
        """ Process menu choices for Keword Example nodes """
        n = evt.GetId() - self.cmd_id_start["KeywordExampleNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)

        if n == 0:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Keyword Example

        elif n == 1:    # Locate this Clip
            # Capture Clip Name
            clipname = self.GetItemText(sel)
            # Load the Clip.  To save time, we can skip the Clip Transcripts.
            tempClip = Clip.Clip(selData.recNum, skipText=True)
            tempCollection = Collection.Collection(tempClip.collection_num)
            collectionList = [tempCollection.id]                     # Initialize a List
            while tempCollection.parent != 0:
                tempCollection = Collection.Collection(tempCollection.parent)
                collectionList.insert(0, tempCollection.id)
            nodeList = [_('Collections')] + collectionList + [clipname]
            self.select_Node(nodeList, 'ClipNode')
            
        elif n == 2:    # Delete
            # For each Keyword Example in the selected items ...
            for item in selItems:
                # ... grab the individual item ...
                sel = item
                # ... and get the item's data
                selData = self.GetPyData(sel)

                kwg = self.GetItemText(self.GetItemParent(self.GetItemParent(sel)))
                kw = self.GetItemText(self.GetItemParent(sel))
                # Get user confirmation of the Keyword Example Add request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Do you want to remove Clip "%s" as an example of Keyword "%s:%s"?'), 'utf8')
                else:
                    prompt = _('Do you want to remove Clip "%s" as an example of Keyword "%s:%s"?')
                dlg = Dialogs.QuestionDialog(self, prompt % (self.GetItemText(sel), kwg, kw))
                result = dlg.LocalShowModal()
                dlg.Destroy()
                # If confirmed ...
                if result == wx.ID_YES:
                    # ... remove keyword example status
                    DBInterface.SetKeywordExampleStatus(kwg, kw, selData.recNum, 0)
                    # Get the Node Data
                    nodeList = ('Keywords', kwg, kw, self.GetItemText(sel))
                    # Call the DB Tree's delete_Node method to remove the keyowrd example node.
                    # Include the Clip Record Number so the correct Clip entry will be removed.
                    self.delete_Node(nodeList, 'KeywordExampleNode', selData.recNum)
            
    def OnSearchCommand(self, event):
        """Handle selections for the Search menu."""
        n = event.GetId() - self.cmd_id_start["SearchRootNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
        
        if n == 0:      # Clear All
            # Delete all child nodes
            self.DeleteChildren(sel)
            # reset the Search Counter
            self.searchCount = 1
            
        elif n == 1:    # Search
            # Process the Search Request
            search = ProcessSearch.ProcessSearch(self, self.searchCount)
            # Get the new searchCount Value (which may or may not be changed)
            self.searchCount = search.GetSearchCount()
            self.EnsureVisible(sel)
            
        else:
            raise MenuIDError

    def OnSearchResultsCommand(self, event):
        """Handle selections for the Search Results menu."""
        n = event.GetId() - self.cmd_id_start["SearchResultsNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]

        if n == 0:    # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                # If data is a list, there are multiple Clip nodes to paste
                if isinstance(data, list):
                    # Iterate through the Clip nodes
                    for datum in data:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                # if data is NOT a list, it is a single Clip node to paste
                else:
                    # ... and paste the data
                    DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 1:      # Clear
            # For each selected item ...
            for sel in selItems:
                # ... delete the node
                self.Delete(sel)
            
        elif n == 2:    # Convert to Collection
            # Get the data associated with the selected item
            sel_data = self.GetPyData(sel)
            # Convert the Search Result to a Collection
            if self.ConvertSearchToCollection(sel, sel_data):
                # Finally, if the Search result is converted, remove the converted Node
                # from the Search Results.
                self.delete_Node((_('Search'), self.GetItemText(sel)), 'SearchResultsNode')

        elif n == 3:    # Search Collection Report
            # Use the Report Generator, passing the selected item as the Search Collection
            # and passing in a pointer to the Tree Control.  We want to show file names, clip time data,
            # Transcripts, and Keywords but not Comments, collection notes, or clip notes by default
            # and include nested collection data by default.
            # (The point of including parameters set to false is that it triggers their availability as options.)
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Collection Report"), 'utf8'),
                                            searchColl=sel,
                                            treeCtrl=self,
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showQuoteText=True,
                                            showTranscripts=True,
                                            showSnapshotImage=0,
                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showCollectionNotes=False,
                                            showQuoteNotes=False,
                                            showClipNotes=False,
                                            showSnapshotNotes=False,
                                            showNested=True,
                                            showHyperlink=True)

        elif n == 4:    # Play All Clips
            # Play All Clips takes the current tree node and the ControlObject as parameters.
            # (The ControlObject is owned not by the _DBTreeCtrl but by its parent)
            PlayAllClips.PlayAllClips(searchColl=sel, controlObject=self.parent.ControlObject, treeCtrl=self)

        elif n == 5:    # Rename
            self.EditLabel(sel)
            
        else:
            raise MenuIDError

    def OnSearchLibraryCommand(self, evt):
        """Handle menu selections for Search Library objects."""
        n = evt.GetId() - self.cmd_id_start["SearchLibraryNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
        
        if n == 0:        # Drop from Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it.
                self.DropSearchResult(sel)
            
        elif n == 1:      # Search Library Report
            # Show the Report Generator.  We pass the tree node as the searchLibrary, as well as
            # passing a reference to the TreeCtrl.  We want to show Keywords.
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Library Report"), 'utf8'),
                                            searchSeries=sel,
                                            treeCtrl=self,
                                            showFile=True,
                                            showTime=True,
                                            showKeywords=True)

        else:
            raise MenuIDError
 
    def OnSearchDocumentCommand(self, evt):
        """Handle menu selections for Search Document objects."""
        n = evt.GetId() - self.cmd_id_start["SearchDocumentNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            document_name = self.GetItemText(sel)
            library_name = self.GetItemText(self.GetItemParent(sel))

        if n == 0:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Document

        elif n == 1:      # Drop for Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it from the search result
                self.DropSearchResult(sel)

        elif n == 2:    # Search Document Report
            # Warn the user that we're using REAL data, now SEARCH data in this report        
            msg = _('Please note that even though you are requesting this report based on Search Results, the data in \nthe report includes only Quotes that are in Collections.  It does not include Search Results Quotes.')
            infodlg = Dialogs.InfoDialog(self.parent, msg)
            infodlg.ShowModal()
            infodlg.Destroy()

            # Use the report Generator.  We pass the REAL Library Name and Document Name, not the Search Node ones
            # because we have no way to get the appropriate Search Node Quotes only.  We want to show file names,
            # quote position data, Quote Text, Keywords but not Comments, or Quote Notes by default.
            # (The point of including parameters set to false is that it triggers their availability as options.)

            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Document Report"), 'utf8'),
                                            seriesName=library_name,
                                            documentName=document_name,
                                            showHyperlink=True,
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showQuoteText=True,
##                                            showSnapshotImage=0,
##                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showQuoteNotes=False ) # ,
##                                            showSnapshotNotes=False)

        elif n == 3:      # Document Keyword Map Report
            self.DocumentKeywordMapReport(selData.recNum, library_name, document_name)
            
        else:
            raise MenuIDError

    def OnSearchEpisodeCommand(self, evt):
        """Handle menu selections for Search Episode objects."""
        n = evt.GetId() - self.cmd_id_start["SearchEpisodeNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            episode_name = self.GetItemText(sel)
            library_name = self.GetItemText(self.GetItemParent(sel))
        
        if n == 0:      # Drop for Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it from the search result
                self.DropSearchResult(sel)

        elif n == 1:    # Search Episode Report
            # Warn the user that we're using REAL data, now SEARCH data in this report        
            msg = _('Please note that even though you are requesting this report based on Search Results, the data in \nthe report includes only Clips that are in Collections.  It does not include Search Results Clips.')
            infodlg = Dialogs.InfoDialog(self.parent, msg)
            infodlg.ShowModal()
            infodlg.Destroy()

            # Use the report Generator.  We pass the REAL Library Name and Episode Name, not the Search Node ones
            # because we have no way to get the appropriate Search Node Clips only.  We want to show file names,
            # clip time data, and Transcripts, Keywords but not Comments, or Clip Notes by default.
            # (The point of including parameters set to false is that it triggers their availability as options.)
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Episode Report"), 'utf8'),
                                            seriesName = library_name,
                                            episodeName = episode_name,
                                            showHyperlink=True,
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showTranscripts=True,
                                            showSnapshotImage=0,
                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showClipNotes=False,
                                            showSnapshotNotes=False)

        elif n == 2:      # Keyword Map Report
            self.EpisodeKeywordMapReport(selData.recNum, library_name, episode_name)
            
        else:
            raise MenuIDError

    def OnSearchTranscriptCommand(self, evt):
        """ Handle menuy selections for Search Transcript menu """
        n = evt.GetId() - self.cmd_id_start["SearchTranscriptNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
        
        if n == 0:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Transcript

        elif n == 1:    # Drop for Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it from the Search Results
                self.DropSearchResult(sel)

        else:
            raise MenuIDError

    def OnSearchCollectionCommand(self, evt):
        """Handle menu selections for Search Collection objects."""
        n = evt.GetId() - self.cmd_id_start["SearchCollectionNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            self.cutCopyInfo['sourceItem'] = sel
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            self.cutCopyInfo['sourceItem'] = sel
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:    # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                # If data is a list, there are multiple Clip nodes to paste
                if isinstance(data, list):
                    # Iterate through the Clip nodes
                    for datum in data:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                # if data is NOT a list, it is a single Clip node to paste
                else:
                    # ... and paste the data
                    DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:    # Drop for Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it from the Search Results
                self.DropSearchResult(sel)

        elif n == 4:    # Search Collection Report
            # Use the Report Generator, passing the selected item as the Search Collection
            # and passing in a pointer to the Tree Control.  We want to show file names, clip time data,
            # Transcripts, and Keywords but not Comments, collection notes, or clip notes by default
            # and include nested collection data by default.
            # (The point of including parameters set to false is that it triggers their availability as options.)
            ReportGenerator.ReportGenerator(controlObject = self.parent.ControlObject,
                                            title=unicode(_("Transana Collection Report"), 'utf8'),
                                            searchColl=sel,
                                            treeCtrl=self,
                                            showFile=True,
                                            showTime=True,
                                            showSourceInfo=True,
                                            showQuoteText=True,
                                            showTranscripts=True,
                                            showSnapshotImage=True,
                                            showSnapshotCoding=True,
                                            showKeywords=True,
                                            showComments=False,
                                            showCollectionNotes=False,
                                            showQuoteNotes=False,
                                            showClipNotes=False,
                                            showSnapshotNotes=False,
                                            showNested=True,
                                            showHyperlink=True)

        elif n == 5:    # Play All Clips
            # Play All Clips takes the current Collection and the ControlObject as parameters.
            # (The ControlObject is owned not by the _DBTreeCtrl but by its parent)
            PlayAllClips.PlayAllClips(searchColl=sel, controlObject=self.parent.ControlObject, treeCtrl=self)

        elif n == 6:    # Rename
            self.EditLabel(sel)
            
        else:
            raise MenuIDError

    def OnSearchQuoteCommand(self, evt):
        """Handle selections for the Search Clip menu."""
        n = evt.GetId() - self.cmd_id_start["SearchQuoteNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            quote_name = self.GetItemText(sel)
            coll_name = self.GetItemText(self.GetItemParent(sel))
            coll_parent_num = self.GetPyData(self.GetItemParent(sel)).parent
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:    # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                # If data is a list, there are multiple Clip nodes to paste
                if isinstance(data, list):
                    # Iterate through the Clip nodes
                    for datum in data:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                # if data is NOT a list, it is a single Clip node to paste
                else:
                    # ... and paste the data
                    DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Clip

        elif n == 4:    # Drop for Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it from the Search Results
                self.DropSearchResult(sel)
            
        elif n == 5:    # Locate Quote in Document
            self.parent.ControlObject.LocateQuoteInDocument(selData.recNum)
            
        elif n == 6:    # Locate Quote in Collection
            # Load the Quote.  To save time, we can skip the Qutoe Text.
            tmpQuote = Quote.Quote(num=selData.recNum, skipText=True)
            # Get the Node Data for the Quote
            nodeList = (_('Collections'),) + tmpQuote.GetNodeData()
            # Highlight that node
            self.select_Node(nodeList, 'QuoteNode')

        elif n == 7:    # Rename
            self.EditLabel(sel)
            
        else:
            raise MenuIDError

    def OnSearchClipCommand(self, evt):
        """Handle selections for the Search Clip menu."""
        n = evt.GetId() - self.cmd_id_start["SearchClipNode"]

        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            clip_name = self.GetItemText(sel)
            coll_name = self.GetItemText(self.GetItemParent(sel))
            coll_parent_num = self.GetPyData(self.GetItemParent(sel)).parent
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:    # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                # If data is a list, there are multiple Clip nodes to paste
                if isinstance(data, list):
                    # Iterate through the Clip nodes
                    for datum in data:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                # if data is NOT a list, it is a single Clip node to paste
                else:
                    # ... and paste the data
                    DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Clip

        elif n == 4:    # Drop for Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it from the Search Results
                self.DropSearchResult(sel)
            
        elif n == 5:    # Locate Clip in Episode
            # Load the Clip.  We DO need the Clip Transcripts here.
            clip = Clip.Clip(selData.recNum)

            # Start exception handling to catch failures due to orphaned clips
            try:
                # Load the Episode
                episode = Episode.Episode(clip.episode_num)

                # If the first transcript has a known source (isn't a known orphan) ...
                if clip.transcripts[0].source_transcript != 0:
                    # ... load the SOURCE transcript for the first Clip Transcript
                    # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                    transcript = Transcript.Transcript(clip.transcripts[0].source_transcript, skipText=True)
                # If the first transcript is a KNOWN orphan ...
                else:
##                    # Get the list of possible replacement source transcripts
##                    transcriptList = DBInterface.list_transcripts(episode.series_id, episode.id)
##                    # If only 1 transcript is in the list ...
##                    if len(transcriptList) == 1:
##                        # ... use that.
##                        transcript = Transcript.Transcript(transcriptList[0][0])
##                    # If there are NO transcripts (perhaps because the Episode is gone, perhaps because it has no Transcripts) ...
##                    elif len(transcriptList) == 0:
##                        # ... raise an exception.  We can't locate a transcript if the Episode is gone or if it has no Transcripts.
##                        raise RecordNotFoundError ('Transcript', 0)
##                    # If there are multiple transcripts to choose from ...
##                    else:
##                        # Initialize a list
##                        strList = []
##                        # Iterate through the list of transcripts ...
##                        for (transcriptNum, transcriptID, episodeNum) in transcriptList:
##                            # ... and extract the Transcript IDs
##                            strList.append(transcriptID)
##                        # Create a dialog where the user can choose one Transcript
##                        dlg = wx.SingleChoiceDialog(self, _('Transana cannot identify the Transcript where this clip originated.\nPlease select the Transcript that was used to create this Clip.'),
##                                                    _('Transana Information'), strList, wx.OK | wx.CANCEL)
##                        # Show the dialog.  If the user presses OK ...
##                        if dlg.ShowModal() == wx.ID_OK:
##                            # ... use the selected transcript
##                            transcript = Transcript.Transcript(dlg.GetStringSelection(), episode.number)
##                        # If the user presses Cancel (Esc, etc.) ...
##                        else:
                            # ... raise an exception
                            raise RecordNotFoundError ('Transcript', 0)
                # As long as a Control Object is defined (which it always will be)
                if self.parent.ControlObject != None:
                    # Set the active transcript to 0 so the whole interface will be reset
                    self.parent.ControlObject.activeTranscript = 0
                    # Load the source Transcript
                    self.parent.ControlObject.LoadTranscript(episode.series_id, episode.id, transcript.id)
                    # We need to signal that the Visualization needs to be re-drawn.
                    self.parent.ControlObject.ChangeVisualization()
                    # For each Clip transcript except the first one (which has already been loaded) ...
                    for tr in clip.transcripts[1:]:
                        # ... load the source Transcript as an Additional Transcript
                        self.parent.ControlObject.OpenAdditionalTranscript(tr.source_transcript)
                    # Mark the Clip as the current selection.  (This needs to be done AFTER all transcripts have been opened.)
                    self.parent.ControlObject.SetVideoSelection(clip.clip_start, clip.clip_stop)

                    # We need the screen to update here, before the next step.
                    wx.Yield()
                    # Now let's go through each Transcript Window ...
                    for trWin in self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren():
                        # ... move the cursor to the TRANSCRIPT's Start Time (not the Clip's)
                        trWin.editor.scroll_to_time(clip.transcripts[trWin.panelNum].clip_start + 10)
                        # .. and select to the TRANSCRIPT's End Time (not the Clip's)
                        trWin.editor.select_find(str(clip.transcripts[trWin.panelNum].clip_stop))

            except:
                (exctype, excvalue, traceback) = sys.exc_info()
                if len(clip.transcripts) == 1:
                    msg = _('The Transcript this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
                else:
                    msg = _('One of the Transcripts this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
                    self.parent.ControlObject.ClearAllWindows()
                dlg = Dialogs.ErrorDialog(self, msg)
                result = dlg.ShowModal()
                dlg.Destroy()

        elif n == 6:    # Locate Clip in Collection
            # Load the Clip.  To save time, we can skip the Clip Transcripts.
            tmpClip = Clip.Clip(selData.recNum, skipText=True)
            # Get the Node Data for the Clip
            nodeList = (_('Collections'),) + tmpClip.GetNodeData()
            # Highlight that node
            self.select_Node(nodeList, 'ClipNode')

        elif n == 7:    # Rename
            self.EditLabel(sel)
            
        else:
            raise MenuIDError

    def OnSearchSnapshotCommand(self, evt):
        """Handle selections for the Search Snapshot menu."""
        
        n = evt.GetId() - self.cmd_id_start["SearchSnapshotNode"]
        # Get the list of selected items
        selItems = self.GetSelections()
        # If there's only one selected item ...
        if len(selItems) == 1:
            # ... grab that item ...
            sel = selItems[0]
            # ... get tthe item's data ...
            selData = self.GetPyData(sel)
            # ... and set some variables
            snapshot_name = self.GetItemText(sel)
            coll_name = self.GetItemText(self.GetItemParent(sel))
            coll_parent_num = self.GetPyData(self.GetItemParent(sel)).parent
        
        if n == 0:      # Cut
            self.cutCopyInfo['action'] = 'Move'    # Functionally, "Cut" is the same as Drag/Drop Move
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 1:    # Copy
            self.cutCopyInfo['action'] = 'Copy'
            # If there's only one selected item ...
            if len(selItems) == 1:
                self.cutCopyInfo['sourceItem'] = sel
            else:
                self.cutCopyInfo['sourceItem'] = selItems
            self.OnCutCopyBeginDrag(evt)

        elif n == 2:    # Paste
            # specify the data formats to accept
            df = wx.CustomDataFormat('DataTreeDragData')
            # Specify the data object to accept data for this format
            cdo = wx.CustomDataObject(df)
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Try to get the appropriate data from the Clipboard      
            success = wx.TheClipboard.GetData(cdo)
            # If we got appropriate data ...
            if success:
                # ... unPickle the data so it's in a usable format
                data = cPickle.loads(cdo.GetData())
                # If data is a list, there are multiple Clip nodes to paste
                if isinstance(data, list):
                    # Iterate through the Clip nodes
                    for datum in data:
                        # ... and paste the data
                        DragAndDropObjects.ProcessPasteDrop(self, datum, sel, self.cutCopyInfo['action'])
                # if data is NOT a list, it is a single Clip node to paste
                else:
                    # ... and paste the data
                    DragAndDropObjects.ProcessPasteDrop(self, data, sel, self.cutCopyInfo['action'])
            # Close the Clipboard
            wx.TheClipboard.Close()

        elif n == 3:      # Open
            self.OnItemActivated(evt)                            # Use the code for double-clicking the Clip

        elif n == 4:    # Drop for Search Results
            # For each selected item ...
            for sel in selItems:
                # ... drop it from the Search Results
                self.DropSearchResult(sel)

        elif n == 5:      # Load Snapshot Context
            # Load the Snapshot.
            snapshot = Snapshot.Snapshot(selData.recNum)
            # If the Snapshot HAS a defined context ...
            if snapshot.transcript_num > 0:
                # Load the Episode
                episode = Episode.Episode(snapshot.episode_num)
                # Initialize a list for Transcript Numbers
                transcriptNums = []
                # Get a list of possible transcripts
                transcripts = DBInterface.list_transcripts(episode.series_id, episode.id)
                # For each Transcript in the Transcript List ...
                for tr in transcripts:
                    # ... add the transcript Number to the Transcript List
                    transcriptNums.append(tr[0])
                # Start exception handling to catch failures due to orphaned clips
                try:
                    # If the transcript is in the Transcript list ...
                    if snapshot.transcript_num in transcriptNums:
                        # ... load the transcript
                        # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
                        transcript = Transcript.Transcript(snapshot.transcript_num, skipText=True)
                    # If any of the transcripts are orphans ...
                    else:
                        # ... raise an exception
                        raise RecordNotFoundError ('Transcript', snapshot.transcript_num)
                    # As long as a Control Object is defined (which it always will be)
                    if self.parent.ControlObject != None:
                        # Set the active transcript to 0 so the whole interface will be reset
                        self.parent.ControlObject.activeTranscript = 0
                        # Load the source Transcript
                        self.parent.ControlObject.LoadTranscript(episode.series_id, episode.id, transcript.id)
                        # Check to see if the load succeeded before continuing!
                        if self.parent.ControlObject.currentObj != None:
                            # We need to signal that the Visualization needs to be re-drawn.
                            self.parent.ControlObject.ChangeVisualization()
                            # Mark the Clip as the current selection.  (This needs to be done AFTER all transcripts have been opened.)
                            self.parent.ControlObject.SetVideoSelection(snapshot.episode_start, snapshot.episode_start + snapshot.episode_duration)
                            # We need the screen to update here, before the next step.
                            wx.Yield()
                            # Now let's go through each Transcript Window ...
                            for trWin in self.parent.ControlObject.TranscriptWindow:
                                # ... move the cursor to the TRANSCRIPT's Start Time (not the Clip's)
                                trWin.dlg.editor.scroll_to_time(snapshot.episode_start + 500)
                                # .. and select to the TRANSCRIPT's End Time (not the Clip's)
                                trWin.dlg.editor.select_find(str(snapshot.episode_start + snapshot.episode_duration))
                                # update the selection text
                                wx.CallLater(50, trWin.dlg.editor.ShowCurrentSelection)
                except:
                    (exctype, excvalue, traceback) = sys.exc_info()

                    if DEBUG:
                        print exctype, excvalue
                        import traceback
                        traceback.print_exc(file=sys.stdout)
                    
                    if len(transcripts) == 1:
                        msg = _('The Transcript this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
                    else:
                        msg = _('One of the Transcripts this Clip was created from cannot be loaded.\nMost likely, the transcript has been deleted.')
                    self.parent.ControlObject.ClearAllWindows()
                    dlg = Dialogs.ErrorDialog(self, msg)
                    result = dlg.ShowModal()
                    dlg.Destroy()
            else:
                msg = unicode(_('Snapshot "%s" does not have a defined context.\nLibrary ID, Episode ID, and Transcript ID must all be selected.'), 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg % snapshot.id)
                result = dlg.ShowModal()
                dlg.Destroy()
                # Load the Snapshot Properties form
                self.parent.edit_snapshot(snapshot)

        elif n == 6:    # Locate Snapshot in Collection
            # Load the Snapshot.
            tmpSnapshot = Snapshot.Snapshot(selData.recNum)
            # Get the Node Data for the Snapshot
            nodeList = (_('Collections'),) + tmpSnapshot.GetNodeData()
            # Highlight that node
            self.select_Node(nodeList, 'SnapshotNode')

        elif n == 7:    # Rename
            self.EditLabel(sel)
            
        else:
            raise MenuIDError

    def CreateMultiTranscriptClip(self, coll_num, coll_name):
        # Begin exception handling
        try:
            # Create a Popup Dialog
            tmpDlg = Dialogs.PopupDialog(None, _("Creating Clip"), _("Gathering Clip Information"))
            # Get the Transcript Selection information from the ControlObject.
            transcriptSelectionInfo = self.parent.ControlObject.GetMultipleTranscriptSelectionInfo()
            # Create a Clip Object
            tempClip = Clip.Clip()
            # If this clip is being created from an Episode record ...
            if type(self.parent.ControlObject.currentObj) == type(Episode.Episode()):
                # ... then the Clip's Episde Number is the Episode's number ... (from the Control Object's current object, the Episode)
                tempClip.episode_num = self.parent.ControlObject.currentObj.number
                # Load the Episode that is connected to the Clip's Originating Transcript
                sourceObj = Episode.Episode(tempClip.episode_num)

            # If this clip ISN'T from an Episode, it's from another Clip ...
            else:
                # ... so we get the new clip's Episode number from the original clip's Episode Number
                tempClip.episode_num = self.parent.ControlObject.currentObj.episode_num
                # Load the Clip we're starting from.  (Get a fresh copy in case KEYWORDS have changed!)
                sourceObj = Clip.Clip(self.parent.ControlObject.currentObj.number)
            # Initialize the clip start time to the end of the media file
            earliestStartTime = self.parent.ControlObject.GetMediaLength(True)
            # Initialise the clip end time to the beginning of the media file
            latestEndTime = 0
            # Initialize the sort order to 0
            sortOrder = 0
            # iterate through the Transcript Selection Info gathered above
            for (transcriptNum, startTime, endTime, text) in transcriptSelectionInfo:
                # If the transcript HAS a selection ...
                if text != "":
                    # ... Create a Transcript object for each Transcript Selection
                    tempTranscript = Transcript.Transcript()
                    # Assign the Transcript's Episode Number
                    tempTranscript.episode_num = tempClip.episode_num
                    # The transcriptNum value is the source transcript information
                    tempTranscript.source_transcript = transcriptNum
                    # The order we iterate through these values matches the sort order
                    tempTranscript.sort_order = sortOrder
                    # Increment the sort order value
                    sortOrder += 1
                    # For Transcript Change Propagation, each CLIP TRANSCRIPT needs to know it's own start and stop times!
                    tempTranscript.clip_start = startTime
                    tempTranscript.clip_stop = endTime
                    # Assign the transcript text to the new transcript
                    tempTranscript.text = text
                    # Add this new transcript to the clip's list of transcripts
                    tempClip.transcripts.append(tempTranscript)
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
            # The new clip's Collection number is the selected record number from the Database Tree
            tempClip.collection_num = coll_num
            # The new clip's Collection ID is the Collection Name, already extracted from the Database Tree
            tempClip.collection_id = coll_name

            # Get the Video Checkbox data to determine what media file(s) to include
            videoCheckboxData = self.parent.ControlObject.GetVideoCheckboxDataForClips(earliestStartTime)
            # If there is no videoCheckbox Data (ie there are no video checkboxes) or the FIRST media files should be included ...
            if (videoCheckboxData == []) or (videoCheckboxData[0][0]):
                # The Clip's Media Filename comes from the Episode Record
                tempClip.media_filename = self.parent.ControlObject.currentObj.media_filename
                # The offset is 0 because the Episode's first video is used
                tempClip.offset = 0
            # audio defaults to 1 (on).  If there are checkboxes and the first audio indicator is unchecked ...
            if (videoCheckboxData != []) and (not videoCheckboxData[0][1]):
                # ... then indicate that the first audio track should not be included.
                tempClip.audio = 0
            # For each set of media player checkboxes after the first (which has already been processed) ...
            for x in range(1, len(videoCheckboxData)):
                # ... get the checkbox data
                (videoCheck, audioCheck) = videoCheckboxData[x]
                # if the media should be included ...
                if videoCheck:
                    # if this is the FIRST included media file, store the data in the Clip object.
                    if tempClip.media_filename == '':
                        tempClip.media_filename = self.parent.ControlObject.currentObj.additional_media_files[x - 1]['filename']
                        tempClip.offset = self.parent.ControlObject.currentObj.additional_media_files[x - 1]['offset']
                        tempClip.audio = audioCheck
                    # If this is NOT the first included media file, store the data in the additional_media_files structure
                    else:
                        tempClip.additional_media_files = {'filename' : self.parent.ControlObject.currentObj.additional_media_files[x - 1]['filename'],
                                                           'length'   : tempClip.clip_stop - tempClip.clip_start,
                                                           'offset'   : self.parent.ControlObject.currentObj.additional_media_files[x - 1]['offset'],
                                                           'audio'    : audioCheck }
            # If NO media files were included, create an error message to that effect.
            if tempClip.media_filename == '':
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Clip Creation cancelled.  No media files have been selected for inclusion.'), 'utf8')
                else:
                    prompt = _('Clip Creation cancelled.  No media files have been selected for inclusion.')
                errordlg = Dialogs.ErrorDialog(None, prompt)
                errordlg.ShowModal()
                errordlg.Destroy()
                # If Clip Creation fails, we don't need to continue any more.
                contin = False
                # Let's get out of here!
                return
            
            # The new clip's start time is the earliest start time from the Transcript Selections.
            tempClip.clip_start = earliestStartTime
            # The new clip's end time is the latest end time from the Trasncript Selections.
            tempClip.clip_stop = latestEndTime

            # We need to determine the Clip's position in the Sort Order
            sel = self.GetSelections()[0]
            selData = self.GetPyData(sel)
            if selData.nodetype == 'CollectionNode':
                # The new clip's Sort Order is the max sort order Value.  (We change it later!)
                tempClip.sort_order = DBInterface.getMaxSortOrder(tempClip.collection_num) + 1
            elif selData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
                tempClip.sort_order = 0

            # We need to set up the initial keywords.
            # Iterate through the source object's Keyword List ...
            for clipKeyword in sourceObj.keyword_list:
                # ... and copy each keyword to the new Clip
                tempClip.add_keyword(clipKeyword.keywordGroup, clipKeyword.keyword)
            # Remove the Popup Dialog
            tmpDlg.Destroy()

            return tempClip
        # Catch exceptions
        except:

            # Remove the Popup Dialog
            tmpDlg.Destroy()

            if DEBUG:
                import traceback
                traceback.print_exc(file=sys.stdout)

            # Display the Exception Message, allow "continue" flag to remain true
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_("Exception %s : %s"), 'utf8')
            else:
                prompt = _("Exception %s : %s")
            errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0], sys.exc_info()[1]))
            errordlg.ShowModal()
            errordlg.Destroy()
            return None

    def DropSearchResult(self, selection):
        """ Remove a Search Result node from the database tree """
        # Get the full Node Branch by climbing it to one level above the root
        nodeList = (self.GetItemText(selection),)
        originalNodeType = self.GetPyData(selection).nodetype
        while (self.GetItemParent(selection) != self.GetRootItem()):
            selection = self.GetItemParent(selection)
            nodeList = (self.GetItemText(selection),) + nodeList
        # Call the DB Tree's delete_Node method.
        self.delete_Node(nodeList, originalNodeType)

    def OnRightDown(self, event):
        """Called when the right mouse button is pressed."""
        # "Cut and Paste" functioning within the DBTree requires that we know information about how the user
        # is interacting with the data tree so that the Paste function can be enabled or disabled.
        # First, let's collect that information to be used later.
        
        # Items in the tree are not automatically selected with a right click.
        # We must select the item that is initially clicked manually!!
        # We do this by looking at the screen point clicked and applying the tree's
        # HitTest method to determine the current item, then actually selecting the item

        # This line works on Windows, but not on Mac or Linux using wxPython 2.4.1.2  due to a problem with event.GetPoint().
        # pt = event.GetPoint()
        # therfore, this alternate method is used.
        # Get the Mouse Position on the Screen in a more generic way to avoid the problem above
#        (windowx, windowy) = wx.GetMousePosition()
        # Translate the Mouse's Screen Position to the Mouse's Control Position
        pt = event.GetPositionTuple()   # self.ScreenToClientXY(windowx, windowy)
        # use HitTest to determine the tree item as the screen point indicated.
        sel_item, flags = self.HitTest(pt)

        try:

            # This platform difference is NOT TRUE of wxPython 2.9.4.0
            # Windows and Mac behave differently.  Windows selects on right-click, while Mac does not.  
            # This code makes them both work the same.
            # So if we're on a Mac and are right-clicking an UNSELECTED item ...
##            if ('wxMac' in wx.PlatformInfo) and \


            # If we are right-clicking an UNSELECTED item ...
            if   (not self.IsSelected(sel_item)):
                # ... unselect ALL items ...
                self.UnselectAll()
                # ... and select the appropriate item in the TreeCtrl
                self.SelectItem(sel_item)
        except wx._core.PyAssertionError, e:
            # Unselect all selected items
            self.UnselectAll()
            sel_item = self.GetRootItem()
        except:
            print "DatabaseTreeTab.OnRightDown() Exception:"
            print sys.exc_info()[0]
            print sys.exc_info()[1]
            import traceback
            traceback.print_exc()
            
        # Check for RIGHT-CLICK MENUS only.  There is a similar section that tests this for drag and drop.
        # IF there are multiple selections in the list, they all need to be the same TYPE of object
        if len(self.GetSelections()) > 1:

            # Quotes, Clips, and Snaps can be handled together
            IsQuoteClipSnap = self.GetPyData(self.GetSelections()[0]).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']
            # Search Quotes, Search Clips, and Search Snaps can be handled together
            IsSearchQuoteClipSnap = self.GetPyData(self.GetSelections()[0]).nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']

            # iterate through the list of selections ...
            for x in self.GetSelections():
               # ... comparing the node type to the reference node type of the selected item ...
               if (IsQuoteClipSnap and not (self.GetPyData(x).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode'])) or \
                  (IsSearchQuoteClipSnap and not (self.GetPyData(x).nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode'])) or \
                  (not (IsQuoteClipSnap) and not (IsSearchQuoteClipSnap) and (self.GetPyData(sel_item).nodetype != self.GetPyData(x).nodetype)):
                    # ... and if they're different, display an error message ...
                    dlg = Dialogs.ErrorDialog(self.parent, _('All selected items must be the same type to manipulate multiple items at once.'))
                    dlg.ShowModal()
                    dlg.Destroy()
                    # ... then get out of here as we cannot proceed.
                    return
        # If there's only a single selection, it's OK to change the selected item to the right-clicked one.
        else:
            # Unselect all selected items
            self.UnselectAll()
            # Select the appropriate item in the TreeCtrl, just to be sure
            self.SelectItem(sel_item)
        # If you click off the tree in the Data Tab, you get an ugly wxPython Assertion Error.  Let's trap that.
        try:
            # Get the Node data associated with the selected item
            sel_item_data = self.GetPyData(sel_item)
            # For Paste to work, we need to remember what item was selected here.  This is because we will have lost
            # what item was selected in the DBTree by the time the user selects "Paste" from the popup menu otherwise.
            self.cutCopyInfo['destItem'] = sel_item

            # We look at the Clipboard here and determine if it has a Transana Data object in it.
            # Then we can poll the item type to determine if the Paste should be enabled.

            # specify the data formats to accept.
            #   Our data could be a DataTreeDragData object if the source is the Database Tree
            dfNode = wx.CustomDataFormat('DataTreeDragData')
            # Open the Clipboard
            wx.TheClipboard.Open()
            # Test to see if one of the custom formats is available.  Otherwise, we get odd error messages
            # on the Mac.
            if wx.TheClipboard.IsSupported(dfNode):
                # Specify the data object to accept data for these formats
                #   A DataTreeDragData object will populate the cdoNode object
                cdoNode = wx.CustomDataObject(dfNode)

                # Create a composite Data Object from the object types defined above
                cdo = wx.DataObjectComposite()
                cdo.Add(cdoNode)

                # Try to get data from the Clipboard
                success = wx.TheClipboard.GetData(cdo)
            # If neither of the custom formats is available ...
            else:
                # ... then we know not to worry about clipboard data being available
                success = False

            # If the data in the clipboard is in an appropriate format ...
            if success:
                # ... unPickle the data so it's in a usable form
                # First, let's try to get the DataTreeDragData object
                try:
                    source_item_data = cPickle.loads(cdoNode.GetData())
                    # If we have a list of items that has only one item ...
                    if isinstance(source_item_data, list) and (len(source_item_data) == 1):
                        # ... we can just use the first one here
                        source_item_data = source_item_data[0]
                except:
                    source_item_data = None
                    # If this fails, that's okay
                    pass

            # If the data in the clipboard is not appropriate ...
            else:
                # ... initialize data to an empty DataTreeDragDropData Object so that comparison operations will work
                source_item_data = DragAndDropObjects.DataTreeDragDropData()
            # Determines which context menu to work with based on what type of item is selected
            if sel_item_data.nodetype in ['LibraryRootNode', 'LibraryNode', 'DocumentNode', 'EpisodeNode', 'TranscriptNode',
                                          'CollectionsRootNode', 'CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode',
                                          'KeywordRootNode', 'KeywordGroupNode', 'KeywordNode', 'KeywordExampleNode', 'NoteNode',
                                          'SearchRootNode', 'SearchResultsNode', 'SearchLibraryNode', 'SearchDocumentNode',
                                          'SearchEpisodeNode', 'SearchTranscriptNode', 'SearchCollectionNode', 'SearchQuoteNode',
                                          'SearchClipNode', 'SearchSnapshotNode']:
                # Set the Context Menu based on the node type
                menu = self.menu[sel_item_data.nodetype]
            # All Note Nodes are treated the same
            elif sel_item_data.nodetype in ['LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode',
                                            'ClipNoteNode', 'SnapshotNoteNode', 'DocumentNoteNode', 'QuoteNoteNode']:
                menu = self.menu["NoteNode"]
            # This should be the root database node
            else:
                menu = self.gen_menu

            # Unless we have a root node type, which is never duplicated, ...
            if not (sel_item_data.nodetype in ['LibraryRootNode', 'CollectionsRootNode', 'KeywordRootNode', 'SearchRootNode']):
                # If there's only one selected item ...
                if len(self.GetSelections()) == 1:
                    # ... enable all menu items by default
                    for x in range(menu.GetMenuItemCount()):
                        menu.FindItemByPosition(x).Enable(True)
                # If there are multiple items selected ...
                else:
                    # ... disable all menu items by default
                    for x in range(menu.GetMenuItemCount()):
                        menu.FindItemByPosition(x).Enable(False)

            # Modify enabled status of individual menu items
            if sel_item_data.nodetype == 'LibraryNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                # If there are multiple items selected ...
                else:
                    # Determine if the Paste menu item should be enabled 
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Library')), True)

            elif  sel_item_data.nodetype == 'DocumentNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # If we're running the Professional version ...
                    if TransanaConstants.proVersion:

                        # We need to know if the document in quesiton is ALREADY OPEN!!
                        # Initialize a list for all open Documents
                        openDocuments = []
                        # For each Transcript Window Notebook Page ...
                        for page in range(self.parent.ControlObject.TranscriptWindow.nb.GetPageCount()):
                            # ... for each Splitter Pane on the Notebook Page ...
                            for pane in self.parent.ControlObject.TranscriptWindow.nb.GetPage(page).GetChildren():
                                # ... if the underlying text is from a Document (not a Transcript) ...
                                if isinstance(pane.editor.TranscriptObj, Document.Document):
                                    # ... then add it to the list of open Documents
                                    openDocuments.append(pane.editor.TranscriptObj.number)
                        
                        # IF the currently loaded object is a DOCUMENT AND
                        # the selected document isn't already loaded in a Transcript Window AND
                        # there are fewer that TransanaConstants.maxTranscriptWindows Transcript Windows already open ...
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Document.Document) and \
                           (not(sel_item_data.recNum in openDocuments)) and \
                           (len(self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren()) < TransanaConstants.maxTranscriptWindows) and \
                           not TransanaConstants.demoVersion:
                            # ... enable the Additional Transcript menu
                            menu.Enable(menu.FindItem(_("Open Additional Document")), True)
                        # if ANY of those conditions fails ...
                        else:
                            # ... disable the Additional Transcript menu
                            menu.Enable(menu.FindItem(_("Open Additional Document")), False)
                # If there are multiple items selected ...
                else:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Determine if the Paste menu item should be enabled 
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # ... disable the Additional Document menu (Too obscure to be worth figuring out the logic here!)
                    menu.Enable(menu.FindItem(_("Open Additional Document")), False)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Document')), True)

            elif  sel_item_data.nodetype == 'EpisodeNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # If we're running the Professional version ...
                    if TransanaConstants.proVersion:
                        # Determine if multiple transcripts option should be enabled
                        trList = DBInterface.list_transcripts(self.GetItemText(self.GetItemParent(sel_item)), self.GetItemText(sel_item))
                        if (len(trList) > 1) and \
                           not TransanaConstants.demoVersion:
                            menu.Enable(menu.FindItem(_("Open Multiple Transcripts")), True)
                        else:
                            menu.Enable(menu.FindItem(_("Open Multiple Transcripts")), False)
                # If there are multiple items selected ...
                else:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Determine if the Paste menu item should be enabled 
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Episode')), True)

            elif sel_item_data.nodetype == 'TranscriptNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # If we're running the Professional version ...
                    if TransanaConstants.proVersion:
                        # the currently loaded object is an Episode Transcript AND
                        # the currently loaded Episode is the same as the selected transcript's parent episode AND
                        # there are fewer that TransanaConstants.maxTranscriptWindows Transcript Windows already open ...
                        if (isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript)) and \
                           (self.parent.ControlObject.TranscriptWindow.GetCurrentObject().clip_num == 0) and \
                           (self.parent.ControlObject.TranscriptWindow.GetCurrentObject().episode_num == sel_item_data.parent) and \
                           (len(self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren()) < TransanaConstants.maxTranscriptWindows) and \
                           not TransanaConstants.demoVersion:
                            # ... enable the Additional Transcript menu
                            menu.Enable(menu.FindItem(_("Open Additional Transcript")), True)
                            # ... but let's check that it's not already loaded!  Iterate through the open Splitter Panes ...
                            for pane in self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren():
                                # ... and check their Transcript Number against the curently selected tree item's Transcript Number.
                                # If they match ...
                                if pane.editor.TranscriptObj.number == sel_item_data.recNum:
                                    # ... disable the Additional Transcript menu
                                    menu.Enable(menu.FindItem(_("Open Additional Transcript")), False)
                                    # ... and stop looking
                                    break
                        # if ANY of those conditions fails ...
                        else:
                            # ... disable the Additional Transcript menu
                            menu.Enable(menu.FindItem(_("Open Additional Transcript")), False)
                # If there are multiple items selected ...
                else:
                    # If a small enough number of transcript are selected ...
                    if len(self.GetSelections()) <= TransanaConstants.maxTranscriptWindows:
                        # ... enable the Open Transcript menu
                        menu.Enable(menu.FindItem(_("Open")), True)
                        # Initialize the Episode indicator
                        episodeNum = None
                        # Check that all of the transcripts are for the same Episode
                        # Iterate through the selected items.
                        for sel_item in self.GetSelections():
                            # Get the data of each selected item
                            sel_item_data = self.GetPyData(sel_item)
                            if episodeNum == None:
                                episodeNum = sel_item_data.parent
                            elif sel_item_data.parent != episodeNum:
                                # ... disable the Additional Transcript menu after all!
                                menu.Enable(menu.FindItem(_("Open")), False)
                                # We can stop looking after one open transcript is found.
                                break
                    # If the currently loaded object is an Episode AND
                    # the currently loaded Episode is the same as the selected transcript's parent episode AND
                    # there fewer than (TransanaConstants.maxTranscriptWindows plus number of selections) Transcript Windows already open ...
                    if (type(self.parent.ControlObject.currentObj) == type(Episode.Episode())) and \
                       ((self.parent.ControlObject.TranscriptWindow.nb.GetPageCount() + len(self.GetSelections())) <= TransanaConstants.maxTranscriptWindows):
                        # ... enable the Additional Transcript menu
                        menu.Enable(menu.FindItem(_("Open Additional Transcript")), True)
                        # Check that none of the selected transcripts is already loaded in a Transcript Window,
                        # and that all of the transcripts are for the right Episode
                        # Iterate through the selected items.
                        for sel_item in self.GetSelections():
                            # Get the data of each selected item
                            sel_item_data = self.GetPyData(sel_item)
                            # If the transcript is already in an open transcript window, or the transcript is from a
                            # different Episode ...
                            if (sel_item_data.recNum in self.parent.ControlObject.TranscriptNum.keys()) or \
                               (self.parent.ControlObject.currentObj.number != sel_item_data.parent):
                                # ... disable the Additional Transcript menu after all!
                                menu.Enable(menu.FindItem(_("Open Additional Transcript")), False)
                                # We can stop looking after one open transcript is found.
                                break
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Transcript')), True)

            elif sel_item_data.nodetype == 'CollectionsRootNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # Unlike most nodes, CollectionsRootNode menu items DON'T get automatically enabled earlier
                    else:
                        menu.Enable(menu.FindItem(_('Paste')), True)

            elif sel_item_data.nodetype == 'CollectionNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # IF the currently loaded object is a DOCUMENT or a Quote ...
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Document.Document) or \
                           isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Quote.Quote):
                            # ... enable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), True)
                        # if ANY of those conditions fails ...
                        else:
                            # ... disable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), False)
                    # Determine if the Add Clip menu item should be enabled
                    if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript):
                        menu.Enable(menu.FindItem(_('Add Clip')), True)
                    else:
                        menu.Enable(menu.FindItem(_('Add Clip')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # Determine if the Add Multi-transcript Clip menu item should be enabled
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript) and \
                           len(self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren()) > 1 and \
                           not TransanaConstants.demoVersion:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), True)
                        else:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), False)
                # If there are multiple items selected ...
                else:
                    # Determine if the Paste menu item should be enabled
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Collection')), True)

            elif sel_item_data.nodetype == 'QuoteNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # IF the currently loaded object is a DOCUMENT or a Quote ...
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Document.Document) or \
                           isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Quote.Quote):
                            # ... enable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), True)
                        # if ANY of those conditions fails ...
                        else:
                            # ... disable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), False)
                    # Determine if the Add Clip menu item should be enabled
                    if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript):
                        menu.Enable(menu.FindItem(_('Add Clip')), True)
                    else:
                        menu.Enable(menu.FindItem(_('Add Clip')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # Determine if the Add Multi-transcript Clip menu item should be enabled
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript) and \
                           len(self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren()) > 1 and \
                           not TransanaConstants.demoVersion:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), True)
                        else:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), False)
                # If there are multiple items selected ...
                else:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Enable Copy menu item
                    menu.Enable(menu.FindItem(_('Copy')), True)
                    # Determine if the Paste menu item should be enabled 
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Items')), True)
                
            elif sel_item_data.nodetype == 'ClipNode':

                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # IF the currently loaded object is a DOCUMENT or a Quote ...
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Document.Document) or \
                           isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Quote.Quote):
                            # ... enable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), True)
                        # if ANY of those conditions fails ...
                        else:
                            # ... disable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), False)
                    # Determine if the Add Clip menu item should be enabled
                    if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript):
                        menu.Enable(menu.FindItem(_('Add Clip')), True)
                    else:
                        menu.Enable(menu.FindItem(_('Add Clip')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # Determine if the Add Multi-transcript Clip menu item should be enabled
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript) and \
                           len(self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren()) > 1 and \
                           not TransanaConstants.demoVersion:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), True)
                        else:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), False)
                        # Load the clip being right-clicked.  To save time, we can skip the Clip Transcripts.
                        tmpClip = Clip.Clip(sel_item_data.recNum, skipText=True)
                        # Get the file extension of the clip's video
                        (nm, ext) = os.path.splitext(tmpClip.media_filename)
                        # Determine if the media format allows "Export Clip Video"
                        # We CANNOT export from multiple simultaneous media files.
                        if (len(tmpClip.additional_media_files) > 0):
                            # ... so disable the export clip video menu option
                            menu.Enable(menu.FindItem(_('Export Clip Video')), False)
                # If there are multiple items selected ...
                else:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Enable Copy menu item
                    menu.Enable(menu.FindItem(_('Copy')), True)
                    # Determine if the Paste menu item should be enabled 
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Items')), True)

            elif sel_item_data.nodetype == 'SnapshotNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # IF the currently loaded object is a DOCUMENT or a Quote ...
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Document.Document) or \
                           isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Quote.Quote):
                            # ... enable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), True)
                        # if ANY of those conditions fails ...
                        else:
                            # ... disable the Add Quote menu item
                            menu.Enable(menu.FindItem(_("Add Quote")), False)
                    # Determine if the Add Clip menu item should be enabled
                    if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript):
                        menu.Enable(menu.FindItem(_('Add Clip')), True)
                    else:
                        menu.Enable(menu.FindItem(_('Add Clip')), False)
                    # If we're running the Professional Version ...
                    if TransanaConstants.proVersion:
                        # Determine if the Add Multi-transcript Clip menu item should be enabled
                        if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript) and \
                           len(self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren()) > 1 and \
                           not TransanaConstants.demoVersion:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), True)
                        else:
                            menu.Enable(menu.FindItem(_('Add Multi-transcript Clip')), False)
                # If there are multiple items selected ...
                else:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Enable Copy menu item
                    menu.Enable(menu.FindItem(_('Copy')), True)
                    # Determine if the Paste menu item should be enabled 
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Items')), True)

            elif sel_item_data.nodetype == 'KeywordGroupNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                # If there are multiple items selected ...
                else:
                    # Determine if the Paste menu item should be enabled
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'KeywordNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'KeywordNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Keyword Group')), True)
                    
            elif sel_item_data.nodetype == 'KeywordNode':
                # If we're running the Professional Version ...
                if TransanaConstants.proVersion:
                    # IF the currently loaded object is a DOCUMENT ...
                    if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Document.Document):
                        # ... enable the Create Quick Quote menu item
                        menu.Enable(menu.FindItem(_("Create Quick Quote")), True)
                    # if ANY of those conditions fails ...
                    else:
                        # ... disable the Create Quick Quote menu item
                        menu.Enable(menu.FindItem(_("Create Quick Quote")), False)
                # Determine if the Add Clip menu item should be enabled
                if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript):
                    menu.Enable(menu.FindItem(_('Create Quick Clip')), True)
                else:
                    menu.Enable(menu.FindItem(_('Create Quick Clip')), False)
                # If we're running the Professional Version ...
                if TransanaConstants.proVersion:
                    # Determine if the Add Multi-transcript Clip menu item should be enabled
                    if isinstance(self.parent.ControlObject.TranscriptWindow.GetCurrentObject(), Transcript.Transcript) and \
                       self.parent.ControlObject.TranscriptWindow.GetCurrentObject().clip_num == 0 and \
                       len(self.parent.ControlObject.TranscriptWindow.nb.GetCurrentPage().GetChildren()) > 1 and \
                       not TransanaConstants.demoVersion:
                        menu.Enable(menu.FindItem(_('Create Multi-transcript Quick Clip')), True)
                    else:
                        menu.Enable(menu.FindItem(_('Create Multi-transcript Quick Clip')), False)
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                # If there are multiple items selected ...
                else:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Enable Copy menu item
                    menu.Enable(menu.FindItem(_('Copy')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Keyword')), True)
                    
            elif sel_item_data.nodetype == 'KeywordExampleNode':
                # If there are multiple items selected ...
                if len(self.GetSelections()) > 1:
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Keyword Example')), True)
                    
            elif sel_item_data.nodetype in ['NoteNode', 'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode', \
                                            'ClipNoteNode', 'SnapshotNoteNode', 'DocumentNoteNode', 'QuoteNoteNode']:
                # If there are multiple items selected ...
                if len(self.GetSelections()) > 1:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Enable Copy menu item
                    menu.Enable(menu.FindItem(_('Copy')), True)
                    # Enable Delete menu item
                    menu.Enable(menu.FindItem(_('Delete Note')), True)
                    
            elif sel_item_data.nodetype == 'SearchRootNode':
                # If Search Results exist ...
                if self.ItemHasChildren(sel_item):
                    # ... enable the "Clear All" menu item
                    menu.Enable(menu.FindItem(_('Clear All')), True)
                # If no Search Resutls exist ...
                else:
                    # ... disable the "Clear All" menu item.  (I could not figure out how to hide it.)
                    menu.Enable(menu.FindItem(_('Clear All')), False)

            elif sel_item_data.nodetype == 'SearchResultsNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                # If there are multiple items selected ...
                else:
                    # Determine if the Paste menu item should be enabled
                    if ((not isinstance(source_item_data, list)) and \
                        (source_item_data.nodetype == 'SearchCollectionNode')) or \
                       ((isinstance(source_item_data, list)) and \
                        (source_item_data[0].nodetype == 'SearchCollectionNode')):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # ... enable the "Clear" menu item
                    menu.Enable(menu.FindItem(_('Clear')), True)
                    
            elif sel_item_data.nodetype == 'SearchLibraryNode':
                # If there are multiple items selected ...
                if len(self.GetSelections()) > 1:
                    # ... enable the "Drop from Search Result" menu item
                    menu.Enable(menu.FindItem(_('Drop from Search Result')), True)
                    
            elif sel_item_data.nodetype == 'SearchDocumentNode':
                # If there are multiple items selected ...
                if len(self.GetSelections()) > 1:
                    # ... enable the "Drop from Search Result" menu item
                    menu.Enable(menu.FindItem(_('Drop from Search Result')), True)
                    
            elif sel_item_data.nodetype == 'SearchEpisodeNode':
                # If there are multiple items selected ...
                if len(self.GetSelections()) > 1:
                    # ... enable the "Drop from Search Result" menu item
                    menu.Enable(menu.FindItem(_('Drop from Search Result')), True)
                    
            elif sel_item_data.nodetype == 'SearchTranscriptNode':
                # If there are multiple items selected ...
                if len(self.GetSelections()) > 1:
                    # ... enable the "Drop from Search Result" menu item
                    menu.Enable(menu.FindItem(_('Drop from Search Result')), True)
                    
            elif sel_item_data.nodetype == 'SearchCollectionNode':
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                # If there are multiple items selected ...
                else:
                    # Determine if the Paste menu item should be enabled
#                    if ((not isinstance(source_item_data, list)) and \
#                        (source_item_data.nodetype == 'SearchCollectionNode')) or \
#                       ((isinstance(source_item_data, list)) and \
#                        (source_item_data[0].nodetype == 'SearchCollectionNode')):
                    if DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), True)
                    # ... enable the "Drop from Search Result" menu item
                    menu.Enable(menu.FindItem(_('Drop from Search Result')), True)
                    
            elif sel_item_data.nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
                # If a single item is selected ...
                if len(self.GetSelections()) == 1:
                    # Determine if the Paste menu item should be enabled 
                    if not DragAndDropObjects.DragDropEvaluation(source_item_data, sel_item_data):
                        menu.Enable(menu.FindItem(_('Paste')), False)
                # If there are multiple items selected ...
                else:
                    # Enable Cut menu item
                    menu.Enable(menu.FindItem(_('Cut')), True)
                    # Enable Copy menu item
                    menu.Enable(menu.FindItem(_('Copy')), True)
                    # ... enable the "Drop from Search Result" menu item
                    menu.Enable(menu.FindItem(_('Drop from Search Result')), True)

            # Close the Clipboard                    
            wx.TheClipboard.Close()

            self.PopupMenu(menu, event.GetPosition())
        except:

            if DEBUG or True:
                print "DatabaseTreeTab.OnRightDown():"
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                import traceback
                traceback.print_exc()
            pass

    def OnItemActivated(self, event):
        """ Handles the double-click selection of items in the Database Tree """
        # If at least one item in the database tree is selected ...
        if len(self.GetSelections()) > 0:
            # Identify the selected item.  Most uses default to the FIRST selected item.
            # (Only multiple transcript selections make sense here.)
            sel_item = self.GetSelections()[0]
            # Get the data associated with the selected item
            sel_item_data = self.GetPyData(sel_item)
            # The ControlObject must be registered for anything to work
            if self.parent.ControlObject != None:

                # if the item is of the appropriate type and is not expanded ...
                if sel_item_data.nodetype in ['LibraryRootNode', 'LibraryNode', 'EpisodeNode', 'CollectionsRootNode', 'CollectionNode', \
                                              'KeywordRootNode', 'KeywordGroupNode', 'SearchLibraryNode', \
                                              'SearchEpisodeNode', 'SearchCollectionNode'] and \
                   self.ItemHasChildren(sel_item) and \
                   not self.IsExpanded(sel_item):

                    # expand it!
                    self.Expand(sel_item)
                        
                # If it's not of the appropriate type, or it's already expanded ...
                else:
                    # If the item is the Library Root, Add a Library
                    if sel_item_data.nodetype == 'LibraryRootNode':
                        # Change the eventID to match "Add Library"
                        event.SetId(self.cmd_id_start["LibraryRootNode"])
                        # Call the Library Root Event Processor
                        self.OnLibraryRootCommand(event)

                    # If the item is a Library, Add an Episode
                    elif sel_item_data.nodetype == 'LibraryNode':
                        # Change the eventID to match "Add Episode"
                        event.SetId(self.cmd_id_start["LibraryNode"] + 1)
                        # Call the Library Event Processor
                        self.OnLibraryCommand(event)
                        
                    # If the item is a Document, load the appropriate objects
                    elif (sel_item_data.nodetype == 'DocumentNode') or (sel_item_data.nodetype == 'SearchDocumentNode'):
                        # Get the list of all selected documents
                        sel_items = self.GetSelections()
                        # Iterate through the selected items
                        for sel_item in sel_items:
                            # Capture the Library Name
                            libraryname=self.GetItemText(self.GetItemParent(sel_item))
                            # Capture the Document Name
                            documentname=self.GetItemText(sel_item)
                            # Load the document via the ControlObject
                            self.parent.ControlObject.LoadDocument(libraryname, documentname, sel_item_data.recNum)

                    # If the item is an Episode, Add a Transcript
                    elif sel_item_data.nodetype == 'EpisodeNode':
                        # Change the eventID to match "Add Transcript"
                        event.SetId(self.cmd_id_start["EpisodeNode"] + 2)
                        # Call the Episode Event Processor
                        self.OnEpisodeCommand(event)
                        
                    # If the item is a Transcript, load the appropriate objects
                    elif (sel_item_data.nodetype == 'TranscriptNode') or (sel_item_data.nodetype == 'SearchTranscriptNode'):
                        # Get the list of all selected transcripts
                        sel_items = self.GetSelections()
                        # Set the Control Object's Active Transcript to 0 to signal the opening of a NEW trascript.
                        self.parent.ControlObject.activeTranscript = 0
                        # Signal that we need to open a first transcript
                        firstTr = True
                        # Iterate through the selected items
                        for sel_item in sel_items:
                            # If we're looking at the first item ...
                            if firstTr:
                                # Capture the Library Name
                                libraryname=self.GetItemText(self.GetItemParent(self.GetItemParent(sel_item)))
                                # Capture the Episode Name
                                episodename=self.GetItemText(self.GetItemParent(sel_item))
                                # Capture the Transcript Name
                                transcriptname=self.GetItemText(sel_item)
                                # Load the first transcript via the ControlObject
                                self.parent.ControlObject.LoadTranscript(libraryname, episodename, transcriptname)
                                # Signal that we now have loaded the first transcript
                                firstTr = False
                            # If we are looking at the second or later items ...
                            else:
                                # If the first Transcript loaded correctly ...
                                if self.parent.ControlObject.activeTranscript != -1:
                                    # Get the item's data
                                    selData = self.GetPyData(sel_item)
                                    # Open the transcript item as an additional Transcript
                                    self.parent.ControlObject.OpenAdditionalTranscript(selData.recNum, libraryname, episodename)

                    # If the item is the Collection Root, Add a Collection
                    elif sel_item_data.nodetype == 'CollectionsRootNode':
                        # Change the eventID to match "Add Collection"
                        event.SetId(self.cmd_id_start['CollectionsRootNode'] + 1)
                        # Call the Collection Root Event Processor
                        self.OnCollRootCommand(event)

                    # If the item is a Collection, Add a Standard Clip / Quote
                    elif sel_item_data.nodetype == 'CollectionNode':
                        # If our current TranscriptWindow Page / Pane is a Transcript ...
                        if self.parent.ControlObject.GetCurrentItemType() == 'Transcript':
                            # ... Change the eventID to match "Add Clip"
                            event.SetId(self.cmd_id_start["CollectionNode"] + 4)
                            # Call the Collection Event Processor
                            self.OnCollectionCommand(event)
                        # If our current TranscriptWindow Page / Pane is a Document or Quote ...
                        elif self.parent.ControlObject.GetCurrentItemType() in ['Document', 'Quote']:
                            # ... Change the eventID to match "Add Quote"
                            event.SetId(self.cmd_id_start["CollectionNode"] + 3)
                            # Call the Collection Event Processor
                            self.OnCollectionCommand(event)

                    # If the item is a Quote, load the appropriate objects
                    elif (sel_item_data.nodetype == 'QuoteNode') or (sel_item_data.nodetype == 'SearchQuoteNode'):
                        # Get the list of all selected documents
                        sel_items = self.GetSelections()
                        # Iterate through the selected items
                        for sel_item in sel_items:
                            # Load the Quote via the ControlObject
                            self.parent.ControlObject.LoadQuote(sel_item_data.recNum)

                    # If the item is a Clip, load the appropriate object.
                    elif (sel_item_data.nodetype == 'ClipNode') or (sel_item_data.nodetype == 'SearchClipNode'):
                        self.parent.ControlObject.LoadClipByNumber(sel_item_data.recNum)  # Load everything via the ControlObject

                    # If the item is a Snapshot, open a Snapshot Window
                    elif (sel_item_data.nodetype == 'SnapshotNode') or (sel_item_data.nodetype == 'SearchSnapshotNode'):
                        # Load the Snapshot that was selected
                        snapshot = Snapshot.Snapshot(sel_item_data.recNum)
                        # Load the Snapshot Interface
                        self.parent.ControlObject.LoadSnapshot(snapshot)    # Load everything via the ControlObject

                    # If the item is the Keyword Root, Add a Keyword Group
                    elif sel_item_data.nodetype == 'KeywordRootNode':
                        # Change the eventID to match "Add Keyword Group"
                        event.SetId(self.cmd_id_start["KeywordRootNode"])
                        # Call the Keyword Root Event Processor
                        self.OnKwRootCommand(event)

                    # If the item is a Keyword Group, Add a Keyword
                    elif sel_item_data.nodetype == 'KeywordGroupNode':
                        # Change the eventID to match "Add Keyword"
                        event.SetId(self.cmd_id_start["KeywordGroupNode"] + 1)
                        # Call the Keyword Group Event Processor
                        self.OnKwGroupCommand(event)

                    # If the item is a Keyword, Create Quick Clip / Quote
                    elif sel_item_data.nodetype == 'KeywordNode':
##                        # If there IS an object loaded in the main interface ...
##                        if self.parent.ControlObject.currentObj != None:
                        # If our current TranscriptWindow Page / Pane is a Transcript ...
                        if self.parent.ControlObject.GetCurrentItemType() in ['Transcript', 'Clip']:
                            if TransanaConstants.proVersion:
                                # ... change the eventID to match "Add Quick Clip"
                                event.SetId(self.cmd_id_start["KeywordNode"] + 5)
                            else:
                                # ... change the eventID to match "Add Quick Clip"
                                event.SetId(self.cmd_id_start["KeywordNode"] + 4)
                            # Call the Keyword Event Processor
                            self.OnKwCommand(event)
                        # If our current TranscriptWindow Page / Pane is a Document or Quote ...
                        elif self.parent.ControlObject.GetCurrentItemType() in ['Document', 'Quote']:
                            # ... change the eventID to match "Add Quick Quote"
                            event.SetId(self.cmd_id_start["KeywordNode"] + 4)
                            # Call the Keyword Event Processor
                            self.OnKwCommand(event)

                    elif sel_item_data.nodetype == 'KeywordExampleNode':
                        # Load the Clip
                        sel_item_data = self.GetPyData(sel_item)                 # Get the Collection's data so we can test its nodetype
                        self.parent.ControlObject.LoadClipByNumber(sel_item_data.recNum)  # Load everything via the ControlObject

                    elif sel_item_data.nodetype in ['NoteNode', 'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', \
                                                    'CollectionNoteNode', 'ClipNoteNode', 'SnapshotNoteNode', 'DocumentNoteNode',
                                                    'QuoteNoteNode']:
                        # We store the record number in the data for the node item
                        num = self.GetPyData(sel_item).recNum
                        # Open the note that was selected
                        n = Note.Note(num)
                        try:
                            # If the NotesBrowser is currently open ...
                            if (self.parent.ControlObject.NotesBrowserWindow != None):
                                # ... make the Notes Browser visible, on top of other windows
                                wx.CallAfter(self.parent.ControlObject.NotesBrowserWindow.Raise)
                                # If the window has been minimized ...
                                if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                                    # ... then restore it to its proper size!
                                    self.parent.ControlObject.NotesBrowserWindow.Iconize(False)
                                # Open the appropriate note
                                self.parent.ControlObject.NotesBrowserWindow.OpenNote(num)
                            # If the Notes Browser is NOT currently open, 
                            else:
                                # ... Obtain a record lock on the note
                                n.lock_record()
                                # Load the note into the Note Editor
                                noteedit = NoteEditor.NoteEditor(self, n.text)
                                # Get User Input
                                n.text = noteedit.get_text()
                                # Save the user's changes to the note
                                n.db_save()
                                # Lock the Note
                                n.unlock_record()
                        # Handle the exception if the record is locked
                        except RecordLockedError, e:
                            self.parent.handle_locked_record(e, _("Note"), n.id)

                    elif sel_item_data.nodetype == 'SearchRootNode':
                        # Process the Search Request
                        search = ProcessSearch.ProcessSearch(self, self.searchCount)
                        # Get the new searchCount Value (which may or may not be changed)
                        self.searchCount = search.GetSearchCount()

    def OnBeginLabelEdit(self, event):
        """ Process a request to edit a Tree Node Label, by vetoing it if it is not a Node that can be edited. """
        # Identify the selected item
        sel_item = event.GetItem()
        # Get the data associated with the selected item
        sel_item_data = self.GetPyData(sel_item)
        # If the selected item is not a Search Result Node
        # or a type that is explicitly handled in OnEndLabelEdit() ...
        if not (sel_item_data.nodetype in ['LibraryNode', 'DocumentNode', 'EpisodeNode', 'TranscriptNode',
                                           'CollectionNode', 'QuoteNode', 'ClipNode', 'SnapshotNode', 'NoteNode',
                                           'LibraryNoteNode', 'EpisodeNoteNode', 'TranscriptNoteNode', 'CollectionNoteNode',
                                           'QuoteNoteNode', 'ClipNoteNode', 'SnapshotNoteNode', 'DocumentNoteNode',
                                           'KeywordNode',
                                           'SearchResultsNode', 
                                           'SearchCollectionNode', 'SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']):
            # ... then veto the edit process before it begins.
            event.Veto()
        # If we're editing a Keyword Node ...
        if (sel_item_data.nodetype == 'KeywordNode'):
            kw_name = self.GetItemText(sel_item)
            kw_group = self.GetItemText(self.GetItemParent(sel_item))
            # We need to update all open Snapshot Windows based on the change in this Keyword
            # Let's make sure the keyword isn't locked in an open Snapshot ...
            if self.parent.ControlObject.UpdateSnapshotWindows('Detect', event, kw_group, kw_name):
                # ... present an error message to the user
                msg = _('Keyword "%s : %s" is contained in a Snapshot you are currently editing.\nYou cannot edit it at this time.')
                msg = unicode(msg, 'utf8')
                dlg = Dialogs.ErrorDialog(self, msg % (kw_group, kw_name))
                dlg.ShowModal()
                dlg.Destroy()
                # ... then veto the edit process before it begins.
                event.Veto()
            
        # If we're editing a Note's label, AND
        #   the Notes Browser is open, AND
        #   the Notes Browser has an open Note, AND
        #   the open note is the one we're trying to edit ...
        if (sel_item_data.nodetype in ['NoteNode', 'LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode', 
                                       'TranscriptNoteNode', 'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode', 
                                       'SnapshotNoteNode']) and \
           ((self.parent.ControlObject.NotesBrowserWindow != None) and \
            (self.parent.ControlObject.NotesBrowserWindow.activeNote != None) and \
            (self.parent.ControlObject.NotesBrowserWindow.activeNote.number == sel_item_data.recNum)):
            # ... then veto the edit process before it begins, since that note is locked
            event.Veto()
            # Raise the Notes Browser Window so it's on top
            self.parent.ControlObject.NotesBrowserWindow.Raise()
            # If the Notes Browser is minimized ...
            if self.parent.ControlObject.NotesBrowserWindow.IsIconized():
                # ... then restore it to its proper size!
                self.parent.ControlObject.NotesBrowserWindow.Iconize(False)

    def OnEndLabelEdit(self, event):
        """ Process the completion of editing a Tree Node Label by altering the underlying Data Object. """
        # Identify the selected item.  Use event.GetItem() rather than self.GetSelection() because self.GetSelection()
        # returns a different value on the Mac if you end a label edit by clicking on a different node in the tree!
        sel_item = event.GetItem()
        # Remember the original name
        originalName = self.GetItemText(sel_item)
        # Get the data associated with the selected item
        sel_item_data = self.GetPyData(sel_item)
        try:
            # If ESC is pressed ...
            if event.IsEditCancelled():
                # ... don't edit the label or do any processing.
                event.Veto()
            # Otherwise ...
            else:
                
                # If we are renaming a Library Record...
                if sel_item_data.nodetype == 'LibraryNode':
                    # Load the Library
                    tempObject = Library.Library(sel_item_data.recNum)
                    # TODO:  MU Messaging needed here!

                # If we are renaming a Document Record...
                elif sel_item_data.nodetype == 'DocumentNode':
                    # Load the Document
                    tempObject = Document.Document(sel_item_data.recNum)
                    
                # If we are renaming an Episode Record...
                elif sel_item_data.nodetype == 'EpisodeNode':
                    # Load the Episode
                    tempObject = Episode.Episode(sel_item_data.recNum)
                    
                # If we are renaming a Transcript Record...
                elif sel_item_data.nodetype == 'TranscriptNode':
                    # Load the Transcript
                    tempObject = Transcript.Transcript(sel_item_data.recNum)
                    
                # If we are renaming a Collection Record...
                elif sel_item_data.nodetype == 'CollectionNode':
                    # Load the Collection
                    tempObject = Collection.Collection(sel_item_data.recNum)

                # If we are renaming a Quote Record...
                elif sel_item_data.nodetype == 'QuoteNode':
                    # Load the Quote.
                    tempObject = Quote.Quote(sel_item_data.recNum)

                # If we are renaming a Clip Record...
                elif sel_item_data.nodetype == 'ClipNode':
                    # Load the Clip.
                    tempObject = Clip.Clip(sel_item_data.recNum)

                # If we are renaming a Snapshot Record...
                elif sel_item_data.nodetype == 'SnapshotNode':
                    # Load the Snapshot.
                    tempObject = Snapshot.Snapshot(sel_item_data.recNum)

                # If we are renaming a Note Record...
                elif sel_item_data.nodetype in ['NoteNode', 'LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode',
                                                'TranscriptNoteNode', 'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode',
                                                'SnapshotNoteNode']:
                    # Load the Note
                    tempObject = Note.Note(sel_item_data.recNum)

                # If we are renaming a Keyword Record...
                elif sel_item_data.nodetype == 'KeywordNode':
                    # Load the Keyword
                    tempObject = Keyword.Keyword(sel_item_data.parent, self.GetItemText(sel_item))
                    # Let's remember the keyword group name.  We may need it later.
                    keywordGroup = self.GetItemText(self.GetItemParent(sel_item))

                # If we are renaming a SearchCollection or a SearchQuote or a SearchClip or a SearchSnapshot ...
                elif sel_item_data.nodetype in ['SearchResultsNode', 'SearchCollectionNode',
                                                'SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']:
                    # ... we don't actually need to do anything, as there is no underlying object that
                    # needs changing.  But we don't veto the Rename either.
                    tempObject = None

                # If we haven't defined how to process the label change, veto it.
                else:
                    # Indicate that no Object has been loaded, so no object will be processed.
                    tempObject = None
                    # Veto the event to cancel the Renaming in the Tree
                    event.Veto()

                # If an object was successfully loaded ...
                if tempObject != None:
                    # We can't just rename a Transcript if it is currently open.  Saving the Transcript that's been
                    # renamed would wipe out the change, and causing problems with data integrity.  Let's test for that
                    # condition and block the rename.
                    if (sel_item_data.nodetype == 'TranscriptNode') and \
                       (tempObject.number in self.parent.ControlObject.TranscriptNum.keys()):
                        dlg = Dialogs.ErrorDialog(self.parent, _('You cannot rename the Transcript that is currently loaded.\nSelect "File" > "New" and try again.'))
                        dlg.ShowModal()
                        dlg.Destroy()
                        event.Veto()
                    elif (sel_item_data.nodetype == 'DocumentNode') and \
                         (self.parent.ControlObject.GetOpenDocumentObject(Document.Document, sel_item_data.recNum) != None):
                            dlg = Dialogs.ErrorDialog(self.parent, _('You cannot rename a Document that is currently loaded.\nClose this Document and try again.'))
                            dlg.ShowModal()
                            dlg.Destroy()
                            event.Veto()
                    elif (sel_item_data.nodetype == 'QuoteNode') and \
                         (self.parent.ControlObject.GetOpenDocumentObject(Quote.Quote, sel_item_data.recNum) != None):
                            dlg = Dialogs.ErrorDialog(self.parent, _('You cannot rename a Quote that is currently loaded.\nClose this Quote and try again.'))
                            dlg.ShowModal()
                            dlg.Destroy()
                            event.Veto()
                    else:
                        # Lock the Object
                        tempObject.lock_record()
                        # If we are renaming a keyword ...
                        if sel_item_data.nodetype == 'KeywordNode':
                            # ... Change the Object's Keyword property
                            tempObject.keyword = Misc.unistrip(event.GetLabel())
                            # We need to remember the new string to see if it has been changed
                            tempStr = tempObject.keyword
                        # If we're not renaming a keyword ...
                        else:
                            # ... Change the Object Name
                            tempObject.id = Misc.unistrip(event.GetLabel())
                            # We need to remember the new string to see if it has been changed
                            tempStr = tempObject.id
                        # Save the Object
                        result = tempObject.db_save()
                        # Unlock the Object
                        tempObject.unlock_record()

                        # If there are leading spaces in the edited label, it causes problems.  The object
                        # doesn't have the leading spaces the label does, so can't be found.  Therefore,
                        # we need to strip() the whitespace, using the unicode-safe unistrip() method.
                        # This could also be triggered by changes to a keyword to correct for use of parentheses.
                        if event.GetLabel() != tempStr:
                            # We have to use CallAfter to change the label we're already in the middle of editing.
#                            wx.CallAfter(self.SetItemText, sel_item, tempStr)
                            wx.CallLater(500, self.SetItemText, sel_item, tempStr)
                        
                        # If we are renaming a Note Record...
                        if sel_item_data.nodetype in ['NoteNode', 'LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode',
                                                      'TranscriptNoteNode', 'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode',
                                                      'SnapshotNoteNode']:
                            # If the Notes Browser is open ...
                            if self.parent.ControlObject.NotesBrowserWindow != None:
                                # Rename the Note in the Notes Browser Tree
                                self.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('R', tempObject, oldName=originalName)

                        # If we're in the Multi-User mode, we need to send a message about the change
                        if not TransanaConstants.singleUserVersion:
                            # Begin constructing the message with the old and new names for the node
                            msg = " >|< %s >|< %s" % (originalName, tempStr)
                            # Get the full Node Branch by climbing it to two levels above the root
                            while (self.GetItemParent(self.GetItemParent(sel_item)) != self.GetRootItem()):
                                # Update the selected node indicator
                                sel_item = self.GetItemParent(sel_item)
                                # Prepend the new Node's name on the Message with the appropriate seperator
                                msg = ' >|< ' + self.GetItemText(sel_item) + msg
                            if 'unicode' in wx.PlatformInfo:
                                libraryPrompt = unicode(_('Libraries'), 'utf8')
                                collectionsPrompt = unicode(_('Collections'), 'utf8')
                                keywordsPrompt = unicode(_('Keywords'), 'utf8')
                            else:
                                # We need the nodeType as the first element.  Then, 
                                # we need the UNTRANSLATED label for the root node to avoid problems in mixed-language environments.
                                libraryPrompt = _('Libraries')
                                collectionsPrompt = _('Collections')
                                keywordsPrompt = _('Keywords')
                            # For Notes, we need to know which root to climb, but we need the UNTRANSLATED root.
                            if self.GetItemText(self.GetItemParent(sel_item)) == libraryPrompt:
                                rootNodeType = 'Libraries'
                            elif self.GetItemText(self.GetItemParent(sel_item)) == collectionsPrompt:
                                rootNodeType = 'Collections'
                            elif self.GetItemText(self.GetItemParent(sel_item)) == keywordsPrompt:
                                rootNodeType = 'Keywords'
                            # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                            # This must be untranslated to avoid problems in mixed-language environments.
                            # Prepend these on the Messsage
                            nodetype = sel_item_data.nodetype
                            # If we have a Note Node, we need to know which kind of Note.
                            if nodetype in ['NoteNode', 'LibraryNoteNode', 'DocumentNoteNode', 'EpisodeNoteNode',
                                            'TranscriptNoteNode', 'CollectionNoteNode', 'QuoteNoteNode', 'ClipNoteNode',
                                            'SnapshotNoteNode']:
                                nodetype = '%sNoteNode' % tempObject.notetype
                            msg = nodetype + " >|< " + rootNodeType + msg
                            if DEBUG:
                                print 'Message to send = "RN %s"' % msg
                            # Send the Rename Node message UNLESS WE ARE DOING KEYWORD MERGE
                            if (TransanaGlobal.chatWindow != None) and ((sel_item_data.nodetype != 'KeywordNode') or (result)):
                                TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)
                            # If we've just renamed a Clip, check for Keyword Examples that need to be renamed
                            if nodetype == 'ClipNode':
                                for (kwg, kw, clipNumber, clipID) in DBInterface.list_all_keyword_examples_for_a_clip(tempObject.number):
                                    nodeList = (_('Keywords'), kwg, kw, originalName)
                                    exampleNode = self.select_Node(nodeList, 'KeywordExampleNode')
                                    self.SetItemText(exampleNode, tempObject.id)
                                    # If we're in the Multi-User mode, we need to send a message about the change
                                    if not TransanaConstants.singleUserVersion:
                                        # Begin constructing the message with the old and new names for the node
                                        msg = " >|< %s >|< %s" % (originalName, tempObject.id)
                                        # Get the full Node Branch by climbing it to two levels above the root
                                        while (self.GetItemParent(self.GetItemParent(exampleNode)) != self.GetRootItem()):
                                            # Update the selected node indicator
                                            exampleNode = self.GetItemParent(exampleNode)
                                            # Prepend the new Node's name on the Message with the appropriate seperator
                                            msg = ' >|< ' + self.GetItemText(exampleNode) + msg
                                        # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                                        # This must be untranslated to avoid problems in mixed-language environments.
                                        # Prepend these on the Messsage
                                        msg = "KeywordExampleNode >|< Keywords" + msg
                                        if DEBUG:
                                            print 'Message to send = "RN %s"' % msg
                                        # Send the Rename Node message
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)

                        # If weve changed a Keyword AND "result" if False, we merged keywords, and therefore need to remove the keyword node!
                        if (sel_item_data.nodetype == 'KeywordNode'):
                            if not result:
                                # Don't actually CHANGE the label when merging keywords.  That would cause a duplicate keyword!
                                event.Veto()
                                # The Mac would crash and burn until I made this a "CallAfter" call
                                wx.CallAfter(self.delete_Node, (_('Keywords'), keywordGroup, originalName), 'KeywordNode')
                            # We need to update all open Snapshot Windows based on the change in this Keyword
                            self.parent.ControlObject.UpdateSnapshotWindows('Update', event, keywordGroup, originalName)
                            # We need to update the Keyword Visualization.  There's no way to tell if the changed
                            # keyword appears or not from here.
                            self.parent.ControlObject.UpdateKeywordVisualization()
                            # Even if this computer doesn't need to update the keyword visualization others, might need to.
                            if not TransanaConstants.singleUserVersion:
                                # We need to update the Keyword Visualization no matter what here, when deleting a keyword group
                                if DEBUG:
                                    print 'Message to send = "UKV %s %s %s"' % ('None', 0, 0)
                                    
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))

                    # Sort the Node
                    # ... get the item's parent
                    tmpNode = self.GetItemParent(sel_item)
                    # ... and sort the parent.  This MUST be in CallAfter, as the node isn't actually renamed yet!!
                    wx.CallAfter(self.SortChildren, tmpNode)

            # Refresh the DB Tree
            self.Refresh()

        # Handle SaveError exceptions (probably raised by Keyword Rename that was cancelled by user when asked about merging.)
        except SaveError:
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
            # If we encounter a SaveError, the error message will have been handled elsewhere.
            # We just need to cancel the Rename event, so the label in the tree will revert to its original value.
            event.Veto()
            # Unlock the Object
            tempObject.unlock_record()

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
            event.Veto()
            dlg = Dialogs.ErrorDialog(self.parent, _("Object Rename failed.  The Object is probably locked by another user."))
            dlg.ShowModal()
            dlg.Destroy()

    def OnKeyDown(self, event):
        """ Handle Key Presses """
        # See if the ControlObject wants to handle the key that was pressed.
        if self.parent.ControlObject.ProcessCommonKeyCommands(event):
            # If so, we're done here.
            return

        # We probably want to call event.Skip()
        callSkip = True
        # If Ctrl is pressed ...
        if event.ControlDown():
            # start exception handling
            try:
                # Get the code of the key pressed
                c = event.GetKeyCode()
                # If we detect the Ctrl key itself ...
                if c == 308:
                    # ... we can skip event.Skip()
                    callSkip = False
                # Ctrl-K is Create Quick Clip
                elif chr(c) == "K":
                    # Create the Quick Clip via the Control Object
                    self.parent.ControlObject.CreateQuickClip()
                    # We can skip event.Skip()
                    callSkip = False
            # If an exception is raised ...
            except:
                # ... we can ignore it as other than a key we need to process
                pass
        # If we should call event.Skip, do so.  This prevents a "beep"
        if callSkip:
            event.Skip()

    def DocumentKeywordMapReport(self, documentNum, libraryName, documentName):
        """ Produce a Keyword Map Report for the specified Library & Document """
        # Create a Keyword Map Report (not embedded)
        frame = KeywordMapClass.KeywordMap(self, -1, _("Transana Document Keyword Map Report"), embedded=False, controlObject = self.parent.ControlObject)
        # Now set it up, passing in the Library and Episode to be displayed
        frame.Setup(documentNum = documentNum, seriesName = libraryName, documentName = documentName)

    def EpisodeKeywordMapReport(self, episodeNum, libraryName, episodeName):
        """ Produce a Keyword Map Report for the specified Library & Episode """
        # Create a Keyword Map Report (not embedded)
        frame = KeywordMapClass.KeywordMap(self, -1, _("Transana Episode Keyword Map Report"), embedded=False, controlObject = self.parent.ControlObject)
        # Now set it up, passing in the Library and Episode to be displayed
        frame.Setup(episodeNum = episodeNum, seriesName = libraryName, episodeName = episodeName)

    def CollectionKeywordMapReport(self, collNum):
        """ Produce a Collection Keyword Map Report for the specified Collection """
        # Create a Keyword Map Report (not embedded)
        frame = KeywordMapClass.KeywordMap(self, -1, _("Transana Collection Keyword Map Report"), embedded=False, controlObject = self.parent.ControlObject)
        # Now set it up, passing in the Library and Episode to be displayed
        frame.Setup(collNum = collNum)

    def KeywordMapLoadItem(self, objType, objNum, ctrlPressed):
        """ This method is called FROM the Keyword Map and causes a specified Quote, Clip, or Snapshot
            to be loaded. """
        if objType == 'Quote':
            if ctrlPressed:
                self.parent.ControlObject.LocateQuoteInDocument(objNum)
            else:
                # Load the specified Quote
                self.parent.ControlObject.LoadQuote(objNum)
            nodeType = 'QuoteNode'
        elif objType == 'Clip':
            if ctrlPressed:
                self.parent.ControlObject.LocateClipInEpisode(objNum)
            else:
                # Load the specified Clip
                self.parent.ControlObject.LoadClipByNumber(objNum)
            nodeType = 'ClipNode'
        elif objType == 'Snapshot':
            # Load the Snapshot that was selected
            snapshot = Snapshot.Snapshot(objNum)
            # Load the Snapshot Interface
            self.parent.ControlObject.LoadSnapshot(snapshot)
        # If Ctrl is NOT pressed, and we have a Quote or Clip, we should select the selected object in the Database Tree.
        if not ctrlPressed and (objType in ['Quote', 'Clip']):
            # Get the Quote or Clip Object
            tempObj = self.parent.ControlObject.currentObj
            # Get the parent Collection
            tempCollection = Collection.Collection(tempObj.collection_num)
            # Add the Collection Root and the Clip to the node list
            nodeList = (_('Collections'), ) + tempCollection.GetNodeData() + (tempObj.id, )
            # Now signal the DB Tree to select / display the selected Clip
            self.select_Node(nodeList, nodeType)
            
    def AnalyticDataExport(self, libraryNum = 0, documentNum = 0, episodeNum = 0, collectionNum = 0):
        """ Implements the Analytic Data Export routine """
        # Create the Analytic Data Export dialog box, passing the appropriate parameter
        if libraryNum > 0:
            clipExport = AnalyticDataExport.AnalyticDataExport(self, -1, libraryNum = libraryNum)
        elif documentNum > 0:
            clipExport = AnalyticDataExport.AnalyticDataExport(self, -1, documentNum = documentNum)
        elif episodeNum > 0:
            clipExport = AnalyticDataExport.AnalyticDataExport(self, -1, episodeNum = episodeNum)
        elif collectionNum > 0:
            clipExport = AnalyticDataExport.AnalyticDataExport(self, -1, collectionNum = collectionNum)
        else:
            clipExport = AnalyticDataExport.AnalyticDataExport(self, -1)
        # Set up the confirmation loop signal variable to get us into the loop
        repeat = True
        # While we are in the confirmation loop ...
        while repeat:
            # ... assume we will want to exit the confirmation loop by default
            repeat = False
            # Get the Analytic Data Export input from the user
            result = clipExport.get_input()
            # if the user clicked OK ...
            if (result != None):
                # ... make sure they entered a file name.
                if result[_('Export Filename')] == '':
                    # If not, create a prompt to inform the user ...
                    prompt = unicode(_('A file name is required'), 'utf8') + '.'
                    # ... and signal that we need to repeat the file prompt
                    repeat = True
                # If they did ...
                else:
                    # ... error check the file name.  If it does not have a PATH ...
                    if os.path.split(result[_('Export Filename')])[0] == u'':
                        # ... add the Video Path to the file name
                        fileName = os.path.join(TransanaGlobal.configData.videoPath, result[_('Export Filename')])
                    # If there is a path, just continue.
                    else:
                        fileName = result[_('Export Filename')]

                    # If the file does not have a .TXT extension ...
                    if fileName[-4:].lower() != '.txt':
                        # ... add one
                        fileName = fileName + '.txt'
                    # Set the FORM's field value to the modified file name
                    clipExport.exportFile.SetValue(fileName)
                    
                    # Check the file name for illegal characters.  First, define illegal characters
                    # (not including PATH characters)
                    illegalChars = '"*?<>|'
                    # For each illegal character ...
                    for char in illegalChars:
                        # ... see if that character appears in the file name the user entered
                        if char in result[_('Export Filename')]:
                            # If so, create a prompt to inform the user ...
                            prompt = unicode(_('There is an illegal character in the file name.'), 'utf8')
                            # ... and signal that we need to repeat the file prompt ...
                            repeat = True
                            # ... and stop looking.
                            break

                # Was there a file name problem or an illegal character?
                if repeat:
                    # If so, display the prompt to inform the user.
                    dlg2 = Dialogs.ErrorDialog(self, prompt)
                    dlg2.ShowModal()
                    dlg2.Destroy()
                    # Signal that we have not gotten a result.
                    result = None
                # If we get to here, check for a duplicate file name
                elif (os.path.exists(fileName)):
                    # If so, create a prompt to inform the user and ask to overwrite the file.
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A file named "%s" already exists.  Do you want to replace it?'), 'utf8')
                    else:
                        prompt = _('A file named "%s" already exists.  Do you want to replace it?')
                    # Create the dialog for the prompt for the user
                    dlg2 = Dialogs.QuestionDialog(None, prompt % fileName)
                    # Center the confirmation dialog on screen
                    dlg2.CentreOnScreen()
                    # Show the confirmation dialog and get the user response.  If the user DOES NOT say to overwrite the file, ...
                    if dlg2.LocalShowModal() != wx.ID_YES:
                        # ... nullify the results of the Analytic Data Export dialog so the file won't be overwritten ...
                        result = None
                        # ... and signal that the user should be re-prompted.
                        repeat = True
                    # Destroy the confirmation dialog
                    dlg2.Destroy()                    
            # If the user didn't press Cancel or decline to overwrite an existing file ...
            if result != None:
                # ... continue with the Analytic Data Export process.
                clipExport.Export()
        # Destroy the Analytic Data Export dialog
        clipExport.Destroy()

    def ConvertSearchToCollection(self, sel, selData):
        """ Converts all the Collections and Clips in a Search Result node to a Collection. """
        # This method takes a tree node and iterates through it's children, coverting them appropriately.
        # If one of those children HAS children of its own, this method should be called recursively.
        
        # print "Converting ", self.GetItemText(sel), selData

        # Things can come up to interrupt our process, but assume for now that we should continue
        contin = True

        # If we have the Search Results Node, create a Collection to be the root of the Converted Search Result
        if selData.nodetype == 'SearchResultsNode':
            # Create a new Collection
            tempCollection = Collection.Collection()
            # Assign the Search Results Name to the Collection
            tempCollection.id = self.GetItemText(sel)
            # Assign the default Owner
            tempCollection.owner = DBInterface.get_username()
            # Load the Collection Properties Form
            contin = self.parent.edit_collection(tempCollection)
            # If the user said OK (did not cancel) ,,,
            if contin:
                # Add the new Collection for the Search Result to the DB Tree
                nodeData = (_('Collections'), tempCollection.id)
                self.add_Node('CollectionNode', nodeData, tempCollection.number, 0, expandNode=True)

                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    msg = "AC %s"
                    data = (nodeData[1],)
                    if DEBUG:
                        print 'Message to send =', msg % data
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage(msg % data)

                # Now that this is a collection, let's update the Node Data to reflect the correct data
                selData.recNum = tempCollection.number
                selData.parent = tempCollection.parent
                # If the user changed the name for the Collection, we need to update the Search Results Node
                self.SetItemText(sel, tempCollection.id)


        if contin:
            # Initialize the Sort Order Counter
            sortOrder = 1
            (childNode, cookieItem) = self.GetFirstChild(sel)
            while childNode.IsOk():
                childData = self.GetPyData(childNode)
                if childData.nodetype == 'SearchLibraryNode':
                    # print "SearchLibrary %s:  Drop this." % self.GetItemText(childNode)
                    pass
                elif childData.nodetype == 'SearchCollectionNode':
                    # print "SearchCollection %s: Convert to Collection." % self.GetItemText(childNode)
            
                    # Load the existing Collection
                    sourceCollection = Collection.Collection(childData.recNum)
                    # Duplicate this Collection
                    newCollection = sourceCollection.duplicate()
                    # The user may have changed the Node Text, indicating they want the new Collection to
                    # have a different Name
                    newCollection.id = self.GetItemText(childNode)
                    # The new Collection has a different parent than the old one!
                    newCollection.parent = selData.recNum
                    # Save the new Collection
                    newCollection.db_save()
                    # Now that this is a Collection, let's update the Node Data to reflect the correct data
                    childData.recNum = newCollection.number
                    childData.parent = newCollection.parent
                    # Add the new Collection for the Search Result to the DB Tree
                    nodeData = ()
                    tempNode = childNode
                    while 1:
                        nodeData = (self.GetItemText(tempNode),) + nodeData
                        tempNode = self.GetItemParent(tempNode)
                        if self.GetPyData(tempNode).nodetype == 'SearchRootNode':
                            break
                        
                    self.add_Node('CollectionNode', (_('Collections'),) + nodeData, newCollection.number, newCollection.parent, expandNode=True)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "AC %s"
                        data = (nodeData[0],)
                        for nd in nodeData[1:]:
                            msg += " >|< %s"
                            data += (nd, )
                        if DEBUG:
                            print 'Message to send =', msg % data
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)


                    if self.ItemHasChildren(childNode):
                        self.ConvertSearchToCollection(childNode, childData)

                elif childData.nodetype == 'SearchQuoteNode':
                    # Load the existing Quote.  We DO need the Quote Text here.
                    sourceQuote = Quote.Quote(childData.recNum)
                    # Duplicate this Quote
                    newQuote = sourceQuote.duplicate()
                    # The user may have changed the Node Text, indicating they want the new Quote to
                    # have a different Name
                    newQuote.id = self.GetItemText(childNode)
                    # The new Quote has a different parent than the old one!
                    newQuote.collection_num = selData.recNum
                    # Add in a Sort Order, which is not carried over during Quote Duplication
                    newQuote.sort_order = sortOrder
                    # Increment the Sort Order Counter
                    sortOrder += 1
                    # Save the new Quote
                    newQuote.db_save()
                    # Now that this is a Quote, let's update the Node Data to reflect the correct data
                    childData.recNum = newQuote.number
                    childData.parent = newQuote.collection_num
                    # Add the new Collection for the Search Result to the DB Tree
                    nodeData = ()
                    tempNode = childNode
                    while 1:
                        nodeData = (self.GetItemText(tempNode),) + nodeData
                        tempNode = self.GetItemParent(tempNode)
                        if self.GetPyData(tempNode).nodetype == 'SearchRootNode':
                            break

                    self.add_Node('QuoteNode', (_('Collections'),) + nodeData, newQuote.number, newQuote.collection_num, sortOrder=newQuote.sort_order, expandNode=True)

                    try:
                        wx.Yield()
                    except:
                        pass

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "AQ %s"
                        data = (nodeData[0],)

                        for nd in nodeData[1:]:
                            msg += " >|< %s"
                            data += (nd, )

                        if DEBUG:
                            print 'Message to send =', msg % data
                            
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                elif childData.nodetype == 'SearchClipNode':

                    # print "SearchClip %s: Convert to Clip." % self.GetItemText(childNode)

                    # Load the existing Clip.  We DO need the Clip Transcripts here.
                    sourceClip = Clip.Clip(childData.recNum)
                    # Duplicate this Clip
                    newClip = sourceClip.duplicate()
                    # The user may have changed the Node Text, indicating they want the new Clip to
                    # have a different Name
                    newClip.id = self.GetItemText(childNode)
                    # The new Clip has a different parent than the old one!
                    newClip.collection_num = selData.recNum
                    # Add in a Sort Order, which is not carried over during Clip Duplication
                    newClip.sort_order = sortOrder
                    # Increment the Sort Order Counter
                    sortOrder += 1
                    # Save the new Clip
                    newClip.db_save()
                    # Now that this is a Clip, let's update the Node Data to reflect the correct data
                    childData.recNum = newClip.number
                    childData.parent = newClip.collection_num
                    # Add the new Collection for the Search Result to the DB Tree
                    nodeData = ()
                    tempNode = childNode
                    while 1:
                        nodeData = (self.GetItemText(tempNode),) + nodeData
                        tempNode = self.GetItemParent(tempNode)
                        if self.GetPyData(tempNode).nodetype == 'SearchRootNode':
                            break

                    self.add_Node('ClipNode', (_('Collections'),) + nodeData, newClip.number, newClip.collection_num, sortOrder=newClip.sort_order, expandNode=True)

                    try:
                        wx.Yield()
                    except:
                        pass

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "ACl %s"
                        data = (nodeData[0],)

                        for nd in nodeData[1:]:
                            msg += " >|< %s"
                            data += (nd, )

                        if DEBUG:
                            print 'Message to send =', msg % data
                            
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                # If we have a Snapshot Node ...
                elif childData.nodetype == 'SearchSnapshotNode':
                    # Load the existing Snapshot.
                    sourceSnapshot = Snapshot.Snapshot(childData.recNum)
                    # Duplicate this Snapshot
                    newSnapshot = sourceSnapshot.duplicate()
                    # The user may have changed the Node Text, indicating they want the new Snapshot to
                    # have a different Name
                    newSnapshot.id = self.GetItemText(childNode)
                    # The new Snapshot has a different parent than the old one!
                    newSnapshot.collection_num = selData.recNum
                    # Add in a Sort Order, which is not carried over during Snapshot Duplication
                    newSnapshot.sort_order = sortOrder
                    # Increment the Sort Order Counter
                    sortOrder += 1
                    # Save the new Snapshot
                    newSnapshot.db_save()
                    # Now that this is a Snapshot, let's update the Node Data to reflect the correct data
                    childData.recNum = newSnapshot.number
                    childData.parent = newSnapshot.collection_num
                    # Add the new Collection for the Search Result to the DB Tree
                    nodeData = ()
                    tempNode = childNode
                    while 1:
                        nodeData = (self.GetItemText(tempNode),) + nodeData
                        tempNode = self.GetItemParent(tempNode)
                        if self.GetPyData(tempNode).nodetype == 'SearchRootNode':
                            break

                    self.add_Node('SnapshotNode', (_('Collections'),) + nodeData, newSnapshot.number, newSnapshot.collection_num, sortOrder=newSnapshot.sort_order, expandNode=True)

                    try:
                        wx.Yield()
                    except:
                        pass

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "ASnap %s"
                        data = (nodeData[0],)

                        for nd in nodeData[1:]:
                            msg += " >|< %s"
                            data += (nd, )

                        if DEBUG:
                            print 'Message to send =', msg % data
                            
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                else:
                    print "DatabaseTreeTab._DBTreeCtrl.ConvertSearchToCollection(): Unhandled Child Node:", self.GetItemText(childNode), childData
                if childNode != self.GetLastChild(sel):
                    (childNode, cookieItem) = self.GetNextChild(sel, cookieItem)
                else:
                    break
        # Return the "contin" value to indicate whether the conversion proceeded.
        return contin

    def GetObjectNodeType(self):
        """ Get the Node Type of the node under the cursor """
        # Get the Mouse Position on the Screen
        (windowx, windowy) = wx.GetMousePosition()
        # Translate the Mouse's Screen Position to the Mouse's Control Position
        (x, y) = self.ScreenToClientXY(windowx, windowy)
        # Now use the tree's HitTest method to find out about the potential drop target for the current mouse position
        (id, flag) = self.HitTest((x, y))

        # Add Exception handling here to handle "off-tree" exception
        try:
            # I'm using GetItemText() here, but could just as easily use GetPyData()
            if self.GetPyData(id) != None:
                destData = self.GetPyData(id).nodetype
            else:
                destData = 'None'
            return destData
        except:
            return 'None'

    def GetSelectedNodeInfo(self):
        """ Get the Name and Type of the currently Selected Node """
        # Get the current selection
        sel = self.GetSelections()
        # Initialize the results to return
        results = []
        # If only a single tree node item is selected ...
        if len(sel) == 1:
            # ... convert from a one-item list to the node itself
            sel = sel[0]
            # Get the Node Data associated with the current selection
            selData = self.GetPyData(sel)
            # Add this item's data to the results list
            results.append((self.GetItemText(sel), selData.recNum, selData.parent, selData.nodetype))
        else:

            IsQuoteClipSnap = self.GetPyData(sel[0]).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']

            # Check for consistency of node type.  As long as node types are consistent, aggregate information about selections.
            for x in sel:
                # ... comparing the node type of the iteration item to the first node type ...
                if (IsQuoteClipSnap and not (self.GetPyData(x).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode'])) or \
                   (not (IsQuoteClipSnap) and (self.GetPyData(sel[0]).nodetype != self.GetPyData(x).nodetype)):
                    # ... and if they're different, display an error message ...
                    dlg = Dialogs.ErrorDialog(self.parent, _('All selected items must be the same type to manipulate multiple items at once.'))
                    dlg.ShowModal()
                    dlg.Destroy()
                    # ... then get out of here as we cannot proceed.
                    return []
                else:
                    # Get the Node Data associated with the current selection
                    selData = self.GetPyData(x)
                    # Add this item's data to the results list
                    results.append((self.GetItemText(x), selData.recNum, selData.parent, selData.nodetype))
        
        # Return the text, record number, parent number, and node type of the current selection
        return results

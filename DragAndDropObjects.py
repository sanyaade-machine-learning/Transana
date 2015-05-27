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

"""This module implements objects that are used in Drag-and-Drop and Cut-andPaste functions.  <BR><BR>

   DragAndDropObjects is made up of the following:<BR><BR>

   DragDropEvaluation(source, destination) -- A Boolean FUNCTION that indicates whether the specified
     source object, a DatabaseTreeTab Node, can be legally dropped on the destination, another
     DatabaseTreeTab Node.  This is implemented as a function because it is used in the DropSource's
     GiveFeedback method and in the DropTarget's OnData method.<BR><BR>

   DataTreeDragDropData -- a Python Object that contains the data that is encapsulated in a
     DatabaseTreeTab Node.  This encapsulates the essential information for what is being dragged.<BR><BR>

   DataTreeDropSource -- an Object derived from wxDropSource.  It is implemented in such a way as to
     provide visual feedback to the user to indicate whether a proposed Drop is legal or not.<BR><BR>

   DataTreeDropTarget -- an Object derived from wxDropTarget.  It is designed to handle data drops on
     the DatabaseTreeTab.<BR><BR>

   ClipDragDropData -- a Python Object that contains the data that is needed for creating a Clip
     by dragging text (and other information) from the Transcript to the Database Tree.<BR><BR>

   ProcessPasteDrop(treeCtrl, sourceData, destNode, action): -- a method that processes a Paste or
     Drop request.  It is a stand-alone method so that as much logic as possible can be shared by
     Drag-and-Drop routines and Cut-and-Paste routines.<BR><BR>
     
   CopyMoveClip(treeCtrl, destNode, sourceClip, sourceCollection, destCollection, action): -- a method
     used by ProcessPasteDrop to implement the copy and move operations for Clips.  It is a stand-alone
     method to allow code reuse between several different methods of copying and moving clips.<BR><BR>
     
   ChangeClipOrder(treeCtrl, destNode, sourceClip, sourceCollection): -- a method used by
     ProcessPasteDrop to implement the altering of Clip Sort Order when desired.  It is a stand-alone
     method to allow code resuse by several different methods related to dropping and pasting clips.

   CreateQuickClip(clipData, kwg, kw, dbTree): -- a method used to implement Open Coding via Quick Clips. """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "DragAndDropObjects DEBUG is ON!!"

import wx                           # Import wxPython
import cPickle                      # use Python's fast cPickle tool instead of the regular Pickle
import sys                          # import Python's sys module

import DBInterface                  # Import Transana's Database Interface
import Library                       # Import the Transana Library object
import Document                     # Import the Transana Document object
import Episode                      # Import the Transana Episode Object
import Transcript                   # Import the Transana Transcript Object
import Collection                   # Import the Transana Collection Object
import Quote                        # Import the Transana Quote Object
import Clip                         # Import the Transana Clip Object
import Snapshot                     # Import the Transana Snapshot Object
import Note                         # Import the Transana Note Object
import ClipPropertiesForm           # Import the Transana Clip Properties Form for adding Clips on Transcript Text Drop
import QuotePropertiesForm          # Import the Transana Quote Properties Form for adding Quotes on Document Text Drop
import KeywordObject as Keyword     # Import the Transana Keyword Object
import KeywordPropertiesForm        # Import the Trasnana Keyword Properties Form for adding Keywords on Transcript Text Drop
import DatabaseTreeTab              # Import the Transana Database Tree Tab Object (for setting _NodeData in manipulating the tree)
import Misc                         # Import the Transana Miscellaneous routines
import Dialogs                      # Import the Transana Dialog Boxes
import TransanaConstants            # Import the Transana Constants
import TransanaGlobal               # Import Transana's Globals
import TransanaExceptions           # Import Transana's Exceptions

# initialize a GLOBAL variable called YESTOALL to handle cross-object communication required for Copy / Move requests
YESTOALL = False

def DragDropEvaluation(source, destination):
    """ This boolean function indicates whether the source tree node can legally be dropped (or pasted) on the destination
        tree node.  This function is encapsulated because it needs to be called from several different locations
        during the Drag-and-Drop process, including the DropSource's GiveFeedback() Method and the DropTarget's
        OnData() Method, as well as the DBTree's OnRightClick() to enable or disable the "Paste" option. """

    if DEBUG:
        print "DragDropEvaluation():"
        print source
        print destination
        print
    
    # If the SOURCE data is not a list but IS a CLIP ...
    if (not isinstance(source, list)) and (source.nodetype == 'ClipNode'):
        # Start exception handling
        try:
            # Try to load the Clip to know that it hasn't been deleted after a COPY.
            # (Trying to paste a clip that's been deleted can trash the database!)
            # Don't load the Clip Transcript to save time.
            tmpClip = Clip.Clip(source.recNum, skipText=True)
        # If the clips does not exist ...
        except:
            # ... then we can't PASTE it, can we?
            return False
    # Return True if the drop is legal, false if it is not.
    # To be legal, we must have a legitimate source and be on a legitimate drop target.
    # If the source is the Database Tree Tab (nodetype = DataTreeDragDropData), then we compare
    # the nodetypes for the source and destination nodes to see if the pairing is compatible.
    # Next, either the record numbers or the nodetype must be different, so you can't drop a node on itself.
    if (source != None) and \
       (destination != None) and \
       (not isinstance(source, ClipDragDropData)) and \
       (not isinstance(source, QuoteDragDropData)) and \
       (not isinstance(source, list)) and \
       ((source.nodetype == 'DocumentNode'         and destination.nodetype == 'LibraryNode'          and source.parent != destination.recNum) or \
        (source.nodetype == 'EpisodeNode'          and destination.nodetype == 'LibraryNode'          and source.parent != destination.recNum) or \
        (source.nodetype == 'CollectionNode'       and destination.nodetype == 'CollectionNode'       and source.parent != destination.recNum) or \
        (source.nodetype == 'CollectionNode'       and destination.nodetype == 'CollectionsRootNode'  and source.parent != 0) or \
        (source.nodetype == 'QuoteNode'            and destination.nodetype == 'CollectionNode'       and source.parent != destination.recNum) or \
        (source.nodetype == 'QuoteNode'            and destination.nodetype == 'QuoteNode') or \
        (source.nodetype == 'QuoteNode'            and destination.nodetype == 'ClipNode') or \
        (source.nodetype == 'QuoteNode'            and destination.nodetype == 'SnapshotNode') or \
        (source.nodetype == 'QuoteNode'            and destination.nodetype == 'KeywordNode') or \
        (source.nodetype == 'ClipNode'             and destination.nodetype == 'CollectionNode'       and source.parent != destination.recNum) or \
        (source.nodetype == 'ClipNode'             and destination.nodetype == 'QuoteNode') or \
        (source.nodetype == 'ClipNode'             and destination.nodetype == 'ClipNode') or \
        (source.nodetype == 'ClipNode'             and destination.nodetype == 'SnapshotNode') or \
        (source.nodetype == 'ClipNode'             and destination.nodetype == 'KeywordNode') or \
        (source.nodetype == 'SnapshotNode'         and destination.nodetype == 'CollectionNode'       and source.parent != destination.recNum) or \
        (source.nodetype == 'SnapshotNode'         and destination.nodetype == 'QuoteNode') or \
        (source.nodetype == 'SnapshotNode'         and destination.nodetype == 'ClipNode') or \
        (source.nodetype == 'SnapshotNode'         and destination.nodetype == 'SnapshotNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'LibraryNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'DocumentNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'EpisodeNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'CollectionNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'QuoteNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'ClipNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'SnapshotNode') or \
        (source.nodetype == 'KeywordNode'          and destination.nodetype == 'KeywordGroupNode') or \
        (source.nodetype == 'LibraryNoteNode'      and destination.nodetype == 'LibraryNode') or \
        (source.nodetype == 'DocumentNoteNode'     and destination.nodetype == 'DocumentNode') or \
        (source.nodetype == 'EpisodeNoteNode'      and destination.nodetype == 'EpisodeNode') or \
        (source.nodetype == 'TranscriptNoteNode'   and destination.nodetype == 'TranscriptNode') or \
        (source.nodetype == 'CollectionNoteNode'   and destination.nodetype == 'CollectionNode') or \
        (source.nodetype == 'QuoteNoteNode'        and destination.nodetype == 'QuoteNode') or \
        (source.nodetype == 'ClipNoteNode'         and destination.nodetype == 'ClipNode') or \
        (source.nodetype == 'SnapshotNoteNode'     and destination.nodetype == 'SnapshotNode') or \
        (source.nodetype == 'SearchCollectionNode' and destination.nodetype == 'SearchResultsNode') or \
        (source.nodetype == 'SearchCollectionNode' and destination.nodetype == 'SearchCollectionNode' and source.parent != destination.recNum) or \
        (source.nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchCollectionNode' and source.parent != destination.recNum) or \
        (source.nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchQuoteNode') or \
        (source.nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchClipNode') or \
        (source.nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchSnapshotNode') or \
        (source.nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchCollectionNode' and source.parent != destination.recNum) or \
        (source.nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchQuoteNode') or \
        (source.nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchClipNode') or \
        (source.nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchSnapshotNode') or \
        (source.nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchCollectionNode' and source.parent != destination.recNum) or \
        (source.nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchQuoteNode') or \
        (source.nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchClipNode') or \
        (source.nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchSnapshotNode')) and \
       ((source.recNum != destination.recNum) or (source.nodetype != destination.nodetype)):
        return True
    # If we have a Clip Creation Object (dragged transcript text, type == ClipDragDropData),
    # or we have a Quote Creation Object (dragged document text, type === QuoteDragDropData),
    # then we can drop it on a Collection or a Quote or a Clip or a Keyword only.
    elif (source != None) and \
         (destination != None) and \
         ((isinstance(source, ClipDragDropData)) or (isinstance(source, QuoteDragDropData))) and \
         ((destination.nodetype == 'CollectionNode') or \
	  (destination.nodetype == 'QuoteNode') or \
	  (destination.nodetype == 'ClipNode') or \
          (destination.nodetype == 'KeywordNode')):
        return True
    # If the source data is a LIST, we have multiple selections!
    elif (source != None) and \
         (destination != None) and \
         (isinstance(source, list)) and \
         ((source[0].nodetype == 'DocumentNode'         and destination.nodetype == 'LibraryNode') or \
          (source[0].nodetype == 'EpisodeNode'          and destination.nodetype == 'LibraryNode') or \
          (source[0].nodetype == 'QuoteNode'            and destination.nodetype == 'CollectionNode') or \
          (source[0].nodetype == 'QuoteNode'            and destination.nodetype == 'QuoteNode') or \
          (source[0].nodetype == 'QuoteNode'            and destination.nodetype == 'ClipNode') or \
          (source[0].nodetype == 'QuoteNode'            and destination.nodetype == 'SnapshotNode') or \
          (source[0].nodetype == 'ClipNode'             and destination.nodetype == 'CollectionNode') or \
          (source[0].nodetype == 'ClipNode'             and destination.nodetype == 'QuoteNode') or \
          (source[0].nodetype == 'ClipNode'             and destination.nodetype == 'ClipNode') or \
          (source[0].nodetype == 'ClipNode'             and destination.nodetype == 'SnapshotNode') or \
          (source[0].nodetype == 'SnapshotNode'         and destination.nodetype == 'CollectionNode') or \
          (source[0].nodetype == 'SnapshotNode'         and destination.nodetype == 'QuoteNode') or \
          (source[0].nodetype == 'SnapshotNode'         and destination.nodetype == 'ClipNode') or \
          (source[0].nodetype == 'SnapshotNode'         and destination.nodetype == 'SnapshotNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'LibraryNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'DocumentNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'EpisodeNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'CollectionNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'QuoteNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'ClipNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'SnapshotNode') or \
          (source[0].nodetype == 'KeywordNode'          and destination.nodetype == 'KeywordGroupNode') or \
          (source[0].nodetype == 'LibraryNoteNode'      and destination.nodetype == 'LibraryNode') or \
          (source[0].nodetype == 'DocumentNoteNode'     and destination.nodetype == 'DocumentNode') or \
          (source[0].nodetype == 'EpisodeNoteNode'      and destination.nodetype == 'EpisodeNode') or \
          (source[0].nodetype == 'TranscriptNoteNode'   and destination.nodetype == 'TranscriptNode') or \
          (source[0].nodetype == 'CollectionNoteNode'   and destination.nodetype == 'CollectionNode') or \
          (source[0].nodetype == 'QuoteNoteNode'        and destination.nodetype == 'QuoteNode') or \
          (source[0].nodetype == 'ClipNoteNode'         and destination.nodetype == 'ClipNode') or \
          (source[0].nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchCollectionNode') or \
          (source[0].nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchQuoteNode') or \
          (source[0].nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchClipNode') or \
          (source[0].nodetype == 'SearchQuoteNode'      and destination.nodetype == 'SearchSnapshotNode') or \
          (source[0].nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchCollectionNode') or \
          (source[0].nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchQuoteNode') or \
          (source[0].nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchClipNode') or \
          (source[0].nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchSnapshotNode') or \
          (source[0].nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchCollectionNode') or \
          (source[0].nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchQuoteNode') or \
          (source[0].nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchClipNode') or \
          (source[0].nodetype == 'SearchSnapshotNode'   and destination.nodetype == 'SearchSnapshotNode')):
        # Assume success
        result = True
        # Iterate through the source list
        for src in source:
            # Check to see if the destination node is IN the source list.  (NODETYPES MUST MATCH TOO!!  recNums aren't enough.)
            if ((src.recNum == destination.recNum) and (source[0].nodetype == destination.nodetype)) and \
                (src.recNum != 0):
                # If so, the evaluation FAILS
                result = False
                # ... and we can stop looking
                break
        # return the result
        return result
    else:
        return False


class DataTreeDragDropData(object):
    """ This is a custom DragDropData object.  It allows the drag to "carry" the information from a node
        from the Database Tree Control. """

    # NOTE:  _NodeType and DataTreeDragDropData have very similar structures so that they can be
    #        used interchangably.  If you alter one, please also alter the other.
   
    def __init__(self, text='', nodetype='Unknown', nodeList=None, recNum=0, parent=0):
        self.text = text           # The source node's text/label
        self.nodetype = nodetype   # The source node's nodetype
        self.nodeList = nodeList   # The Source Node's nodeList (for SearchCollectionNode, SearchQuoteNode, SearchClipNode, and SearchSnapshotNode Cut and Paste)
        self.recNum = recNum       # the source node's record number
        self.parent = parent       # The source node's parent's record number (or Keyword Group name, if the node is a Keyword)

    def __repr__(self):
        """ Return a String Representation of the contents of the DataTreeDragDrop Object """
        str = 'Node %s of type %s, recNum %s, parent %s' % (self.text, self.nodetype, self.recNum, self.parent)
        if self.nodeList != None:
            str = str + '\nnodeList = %s' % (self.nodeList,)
        return str


class DataTreeDropSource(wx.DropSource):
    """ This is a custom DropSource object designed to drag objects from the Data Tree tab and to
        provide feedback to the user during the drag. """
    def __init__(self, tree):
        # Create a Standard wxDropSource Object
        wx.DropSource.__init__(self, tree)
        # Remember the control that initiate the Drag for later use
        self.tree = tree

    # SetData accepts an object (obj) that has been prepared for the DropSource SetData() method
    # by being put into a wxCustomDataObject
    def SetData(self, obj):
        # Set the prepared object as the wxDropSource Data
        wx.DropSource.SetData(self, obj)
        # hold onto the original data, in a usable form, for later use
        self.data = cPickle.loads(obj.GetData())

    def InTranscript(self, windowx, windowy):
        """Determine if the given X/Y position is within the Transcript editor."""
        (transLeft, transTop, transWidth, transHeight) = self.tree.parent.ControlObject.GetTranscriptDims()
        transRight = transLeft + transWidth
        transBot = transTop + transHeight
        return (windowx >= transLeft and windowx <= transRight and windowy >= transTop and windowy <= transBot)

    # I want to provide the user with feedback about whether their drop will work or not.
    def GiveFeedback(self, effect):
        # NOTE:  This method was generating an exception when moving off the data tree.  Thus, exception handling
        #        was added.
        try:
            # This method does not provide the x, y coordinates of the mouse within the control, so we
            # have to figure that out the hard way. (Contrast with DropTarget's OnDrop and OnDragOver methods)
            # Get the Mouse Position on the Screen
            (windowx, windowy) = wx.GetMousePosition()

            if self.InTranscript(windowx, windowy):
                if self.data.nodetype == 'KeywordNode':
                    # Make sure the cursor reflects an acceptable drop.  (This resets it if it was previously changed
                    # to indicate a bad drop.)
                    self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                    # FALSE indicates that feedback is NOT being overridden, and thus that the drop is GOOD!
                    return False
                else:
                    # Set the cursor to give visual feedback that the drop will fail.
                    self.tree.SetCursor(wx.StockCursor(wx.CURSOR_NO_ENTRY))
                    # Setting the Effect to wxDragNone has absolutely no effect on the drop, if I understand this correctly.
                    effect = wx.DragNone
                    # returning TRUE indicates that the default feedback IS being overridden, thus that the drop is BAD!
                    return True

            # Translate the Mouse's Screen Position to the Mouse's Control Position
            (x, y) = self.tree.ScreenToClientXY(windowx, windowy)
            # Now use the tree's HitTest method to find out about the potential drop target for the current mouse position
            (id, flag) = self.tree.HitTest((x, y))
            # I'm using GetItemText() here, but could just as easily use GetPyData()
            destData = self.tree.GetPyData(id)

            # See if we need to scroll the database tree up or down here.  (DatabaseTreeTab.OnMotion used to handle this, but
            # that method no longer gets called during a Drag.)
            (w, h) = self.tree.GetClientSizeTuple()

            # If we are dragging at the top of the window, scroll down
            if y < 8:
              
                # The wxWindow.ScrollLines() method is only implemented on Windows.  We must use something different on the Mac.
                if "wxMSW" in wx.PlatformInfo:
                    self.tree.ScrollLines(-2)
                else:
                    # Suggested by Robin Dunn
                    first = self.tree.GetFirstVisibleItem()
                    prev = self.tree.GetPrevSibling(first)
                    if prev:
                        # drill down to find last expanded child
                        while self.tree.IsExpanded(prev):
                            prev = self.tree.GetLastChild(prev)
                    else:
                        # if no previous sub then try the parent
                        prev = self.tree.GetItemParent(first)

                    if prev:
                        self.tree.ScrollTo(prev)
                    else:
                        self.tree.EnsureVisible(first)
              
            # If we are dragging at the bottom of the window, scroll up
            elif y > h - 8:
                 # The wxWindow.ScrollLines() method is only implemented on Windows.  We must use something different on the Mac.
                 if "wxMSW" in wx.PlatformInfo:
                     self.tree.ScrollLines(2)
                 else:
                     # Suggested by Robin Dunn
                     # first find last visible item by starting with the first
                     next = None
                     last = None
                     item = self.tree.GetFirstVisibleItem()
                     while item:
                         if not self.tree.IsVisible(item): break
                         last = item
                         item = self.tree.GetNextVisible(item)

                     # figure out what the next visible item should be,
                     # either the first child, the next sibling, or the
                     # parent's sibling
                     if last:
                         if self.tree.IsExpanded(last):
                             next = self.tree.GetFirstChild(last)[0]
                         else:
                             next = self.tree.GetNextSibling(last)
                             if not next:
                                 prnt = self.tree.GetItemParent(last)
                                 if prnt:
                                     next = self.tree.GetNextSibling(prnt)

                     if next:
                         self.tree.ScrollTo(next)
                     elif last:
                         self.tree.EnsureVisible(last)

            # This line compares the data being dragged (self.data) to the potential drop site given by the current
            # mouse position (destData).  If the drop is legal,
            # we return FALSE to indicate that we should use the default drag-and-drop feedback, which will indicate
            # that the drop is legal.  If not, we return TRUE to indicate we are using our own feedback, which is
            # implemented by changing the cursor to a "No_Entry" cursor to indicate the drop is not allowed.
            # Note that this code here does not prevent the drop.  That has to be implemented in the Drop Target
            # object.  It just provides visual feedback to the user.  The same evaluatoin function is called elsewhere 
            # (in OnData) when the drop is actually processed.
            if DragDropEvaluation(self.data, destData):
                # Make sure the cursor reflects an acceptable drop.  (This resets it if it was previously changed
                # to indicate a bad drop.)
                self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                # FALSE indicates that feedback is NOT being overridden, and thus that the drop is GOOD!
                return False
            else:
                # Set the cursor to give visual feedback that the drop will fail.
                self.tree.SetCursor(wx.StockCursor(wx.CURSOR_NO_ENTRY))
                # Setting the Effect to wxDragNone has absolutely no effect on the drop, if I understand this correctly.
                effect = wx.DragNone
                # returning TRUE indicates that the default feedback IS being overridden, thus that the drop is BAD!
                return True
        except:

            if DEBUG:
                print "DragAndDropObjects.GiveFeedback(): Exception!!"
                print sys.exc_info()[0]
                print sys.exc_info()[1]
                import traceback
                traceback.print_exc(file=sys.stdout)
                print
            
            # We don't need anything fancy here.  If there's a problem, it's not a valid drop, that's all.
            # Set the cursor to give visual feedback that the drop will fail.
            self.tree.SetCursor(wx.StockCursor(wx.CURSOR_NO_ENTRY))
            # Setting the Effect to wxDragNone has absolutely no effect on the drop, if I understand this correctly.
            effect = wx.DragNone
            # returning TRUE indicates that the default feedback IS being overridden, thus that the drop is BAD!
            return True


class DataTreeDropTarget(wx.PyDropTarget):
    """ This is a custom DropTarget object designed to match drop behavior to the feedback given by the custom
        Drag Object's GiveFeedback() method. """
    def __init__(self, tree):
        # use a normal wxPyDropTarget
        wx.PyDropTarget.__init__(self)
        # Remember the source Tree Control for later use
        self.tree = tree

        # specify the data format to accept Data from Tree Nodes
        self.dfNode = wx.CustomDataFormat('DataTreeDragData')
        # Specify the data object to accept data for this format
        self.sourceNodeData = wx.CustomDataObject(self.dfNode)

        # specify the data format to accept Data from Transcripts to create Clips
        self.dfClip = wx.CustomDataFormat('ClipDragDropData')
        # Specify the data object to accept data for this format
        self.clipData = wx.CustomDataObject(self.dfClip)

        # specify the data format to accept Data from Documents to create Quotes
        self.dfQuote = wx.CustomDataFormat('QuoteDragDropData')
        # Specify the data object to accept data for this format
        self.quoteData = wx.CustomDataObject(self.dfQuote)

        # Create a Composite Data Object
        self.doc = wx.DataObjectComposite()
        # Add the Tree Node Data Object
        self.doc.Add(self.sourceNodeData)
        # Add the Clip Data Object
        self.doc.Add(self.clipData)
        # Add the Quote Data Object
        self.doc.Add(self.quoteData)

        # Set the Composite Data object defined above as the DataObject for the PyDropTarget
        self.SetDataObject(self.doc)

        # Now let's put empty objects in both parts of the wx.DataObjectComposite, so that
        # the OnData logic doesn't blow up when it tries to sort out what's been dropped.
        # (Everything worked OK on Win2K without this, but not on WinXP.)
        self.ClearSourceNodeData()
        self.ClearClipData()

    def ClearSourceNodeData(self):
        """ Clears Data from the Tree Node Data Object """
        # Create a blank Tree Node Data object
        tempData = DataTreeDragDropData()
        # Pickle it
        pickledTempData = cPickle.dumps(tempData, 1)
        # Replace the old Tree Node Data Object with the new empty one
        self.sourceNodeData.SetData(pickledTempData)

    def ClearClipData(self):
        """ Clears Data from the Clip Creation Data Object """
        # Create a blank Clip Creation Data object
        tempData = ClipDragDropData()
        # Pickle it
        pickledTempData = cPickle.dumps(tempData, 1)
        # Replace the old Clip Creation Data Object with the new empty one
        self.clipData.SetData(pickledTempData)
        # Create a blank Quote Creation Data object
        tempData2 = QuoteDragDropData()
        # Pickle it
        pickledTempData2 = cPickle.dumps(tempData2, 1)
        # Replace the old Quote Creation Data Object with the new empty one
        self.quoteData.SetData(pickledTempData2)
        

    def OnEnter(self, x, y, dragResult):
        # Just allow the normal wxDragResult to pass through here
        return dragResult

    def OnLeave(self):
        pass

    def OnDrop(self, x, y):
        # Process the "Drop" event
        # If you drop off the Database Tree, you get an exception here
        try:
            # Use the tree's HitTest method to find out about the potential drop target for the current mouse position
            (self.dropNode, flag) = self.tree.HitTest((x, y))
            # Remember the Drop Location for later Processing (in OnData())
            self.dropData = self.tree.GetPyData(self.dropNode)
            # We don't yet have enough information to veto the drop, so return TRUE to indicate
            # that we should proceed to the OnData method
            return True
        except:
            # If an exception is raised, Veto the drop as there is no Drop Target.
            return False

    def OnData(self, x, y, dragResult):
        # once OnDrop returns TRUE, this method is automatically called.
        global YESTOALL
        YESTOALL = False

        # Let's get the data being dropped so we can do some processing logic
        if self.GetData():
            # First, extract the actual data passed in by the DataTreeDropSource, which used cPickle to pack it.

            try:
                # Try to unPickle the Tree Node Data Object.  If the first Drag is for Clip Creation, this
                # will raise an exception.  If there is a good Tree Node Data Object being dragged, or if
                # one from a previous drag has been Cleared, this will be successful.
                sourceDataList = cPickle.loads(self.sourceNodeData.GetData())
                # This line compares the data being dragged (sourceData) to the drop site determined in OnDrop and
                # passed here as self.dropData.
                if DragDropEvaluation(sourceDataList, self.dropData):
                    # if the sourceDataList is NOT a list (i.e. a single tree node item instead of mulitple selections) ...
                    if not isinstance(sourceDataList, list):
                        # ... then make it into a list so we can iterate through it
                        sourceDataList = [sourceDataList]
                    # if there's only one item being dropped and we're dealing with Keywords, we want confirmation dialogs
                    # from the DropKeyword() method
                    confirmations = (len(sourceDataList) == 1) and (sourceDataList[0].nodetype == 'KeywordNode')
                    # If we DON'T want confirmation dialogs from within DropKeyword(), we want a single confirmation up front,
                    # BUT ONLY IF WE ARE DROPPING KEYWORDS!!  AND WE'RE NOT DROPPING ON A KEYWORD GROUP NODE!!!
                    if not confirmations and (sourceDataList[0].nodetype == 'KeywordNode') and (self.dropData.nodetype != 'KeywordGroupNode'):
                        # Prepare the prompt information
                        if self.dropData.nodetype == 'LibraryNode':
                            # Get user confirmation of the Keyword Add request
                            prompt = unicode(_('Do you want to add multiple Keywords to all %s in %s "%s"?'), 'utf8')
                            data1 = unicode(_('Items'), 'utf8')
                            data2 = unicode(_('Library'), 'utf8')
                            data = (data1, data2, self.tree.GetItemText(self.dropNode))
                        elif self.dropData.nodetype == 'DocumentNode':
                            # Get user confirmation of the Keyword Add request
                            prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                            data1 = unicode(_('Document'), 'utf8')
                            data = (data1, self.tree.GetItemText(self.dropNode))
                        elif self.dropData.nodetype == 'EpisodeNode':
                            # Get user confirmation of the Keyword Add request
                            prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                            data1 = unicode(_('Episode'), 'utf8')
                            data = (data1, self.tree.GetItemText(self.dropNode))
                        elif self.dropData.nodetype == 'CollectionNode':
                            # Get user confirmation of the Keyword Add request
                            prompt = unicode(_('Do you want to add multiple Keywords to all %s in %s "%s"?'), 'utf8')
                            data1 = unicode(_('Items'), 'utf8')
                            data2 = unicode(_('Collection'), 'utf8')
                            data = (data1, data2, self.tree.GetItemText(self.dropNode))
                        elif self.dropData.nodetype == 'QuoteNode':
                            # Get user confirmation of the Keyword Add request
                            prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                            data1 = unicode(_('Quote'), 'utf8')
                            data = (data1, self.tree.GetItemText(self.dropNode))
                        elif self.dropData.nodetype == 'ClipNode':
                            # Get user confirmation of the Keyword Add request
                            prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                            data1 = unicode(_('Clip'), 'utf8')
                            data = (data1, self.tree.GetItemText(self.dropNode))
                        elif self.dropData.nodetype == 'SnapshotNode':
                            # Get user confirmation of the Keyword Add request
                            prompt = unicode(_('Do you want to add multiple Keywords to %s "%s"?'), 'utf8')
                            data1 = unicode(_('Snapshot'), 'utf8')
                            data = (data1, self.tree.GetItemText(self.dropNode))
                        else:
                            prompt = unicode('DataTreeDropTarget.OnData():  Unknown dropData.nodetype.\nPlease press "No".', 'utf8')
                            data = ()
                        # Display the prompt for user feedback
                        dlg = Dialogs.QuestionDialog(None, prompt % data)
                        result = dlg.LocalShowModal()
                        dlg.Destroy()
                    # If we will be collecting confirmation later, ...
                    else:
                        # ... act as if the user pressed Yes to be able to continue
                        result = wx.ID_YES
                    # If the user said yes, or we didn't ask anything ...
                    if result == wx.ID_YES:
                        # If we have multiple Keywords dropped on a Library Node ...
                        if (len(sourceDataList) > 1) and \
                           (sourceDataList[0].nodetype == 'KeywordNode') and \
                           (self.dropData.nodetype in ['LibraryNode', 'EpisodeNode', 'DocumentNode']):

                            # Create a Keyword List
                            kwList = []

                            # Iterate through the source data list
                            for sourceData in sourceDataList:
                                # Create a temporary keyword
                                tmpKeyword = Keyword.Keyword(sourceData.parent, sourceData.text)
                                # Append the keyword to the Keyword List
                                kwList.append(tmpKeyword)
                            # Start handling Exceptions
                            try:
                                # If we're dropping on a Library ...
                                if self.dropData.nodetype == 'LibraryNode':
                                    # Load the dropped-on Library
                                    tmpLibrary = Library.Library(self.dropData.recNum)
                                    # Now get a list of all Documents in the Library and iterate through them
                                    for tempDocumentNum, tempDocumentID, tempLibraryNum in DBInterface.list_of_documents(tmpLibrary.number):
                                        # ... propagating the new Document Keywords to all Quotes from that Document
                                        TransanaGlobal.menuWindow.ControlObject.PropagateObjectKeywords(_('Document'), tempDocumentNum, kwList)
                                    # Now get a list of all Episodes in the Library and iterate through them
                                    for tempEpisodeNum, tempEpisodeID, tempLibraryNum in DBInterface.list_of_episodes_for_series(tmpLibrary.id):
                                        # Propagate the Keyword List to each Episode in the Library
                                        TransanaGlobal.menuWindow.ControlObject.PropagateObjectKeywords(_('Episode'), tempEpisodeNum, kwList)
                                # If we're dropping on a Document ...
                                elif self.dropData.nodetype == 'DocumentNode':
                                    # Propagate the Keyword List to the Document Quotes
                                    TransanaGlobal.menuWindow.ControlObject.PropagateObjectKeywords(_('Document'), self.dropData.recNum, kwList)
                                # If we're dropping on an Episode ...
                                elif self.dropData.nodetype == 'EpisodeNode':
                                    # Propagate the Keyword List to the Episode Clips
                                    TransanaGlobal.menuWindow.ControlObject.PropagateObjectKeywords(_('Episode'), self.dropData.recNum, kwList)
                            # If an exception arises ...
                            except:
                                # ... add the exception to the error log
                                print "EXCEPTION:"
                                print sys.exc_info()[0]
                                print sys.exc_info()[1]
                                import traceback
                                traceback.print_exc(file=sys.stdout)
                        # Iterate through the source data list
                        for sourceData in sourceDataList:
                            # If a previous drag of a Tree Node Data Object has been cleared, the sourceData.nodetype
                            # will be "Unknown", which indicated that the current Drag is NOT a node from the Database
                            # Tree Tab, and therefore should be processed elsewhere.  If it is NOT "Unknown", we should
                            # process it here.  The Type comparison was added to get this working on the Mac.
                            if (type(sourceData) == type(DataTreeDragDropData())) and \
                               (sourceData.nodetype != 'Unknown'):
                                # If we meet the criteria, we process the drop.  We do that here because we have full
                                # knowledge of the Dragged Data and the Drop Target's data here and nowhere else.
                                # Determine if we're copying or moving data.  (In some instances, the 'action' is ignored.)
                                if dragResult == wx.DragCopy:
                                    ProcessPasteDrop(self.tree, sourceData, self.dropNode, 'Copy', confirmations=confirmations)
                                elif dragResult == wx.DragMove:
                                    ProcessPasteDrop(self.tree, sourceData, self.dropNode, 'Move', confirmations=confirmations)
                            else:
                                # If the DragDropEvaluation() test fails, we prevent the drop process by altering the wxDropResult (dragResult)
                                dragResult = wx.DragNone
                    else:
                        # If the DragDropEvaluation() test fails, we prevent the drop process by altering the wxDropResult (dragResult)
                        dragResult = wx.DragNone
                        
                # Once the drop is done or rejected, we must clear the Tree Node data out of the DropTarget.
                # If we don't, this data will still be there if a Clip drag occurs, and there is no way in that
                # circumstance to know which of the dragged objects to process!  Clearing avoids that problem.
                self.ClearSourceNodeData()
                # Reset the cursor, regardless of whether the drop succeeded or failed.
                self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

            except:
                # If an expection occurs here, it's no big deal.  Forget about it.
                # Reset the cursor, regardless of whether the drop succeeded or failed.
                self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                pass

            try:
                # Try to unPickle the Clip Creation Data Object.  If the first Drag is from the Database Tree, this
                # will raise an exception.  If there is a good Clip Creation Data Object being dragged, or if
                # one from a previous drag has been Cleared, this will be successful.
                clipData = cPickle.loads(self.clipData.GetData())
                quoteData = cPickle.loads(self.quoteData.GetData())

                # Dropping Transcript Text onto a Collection, Quote, Clip, or Snapshot creates a Regular Clip.
                # See if the Drop Target is the correct Node Type.  The type comparison was added to get this working on the Mac.
                if (isinstance(clipData, ClipDragDropData)) and \
                   ((self.dropData.nodetype == 'CollectionNode') or \
                    (self.dropData.nodetype == 'QuoteNode') or \
                    (self.dropData.nodetype == 'ClipNode') or \
                    (self.dropData.nodetype == 'SnapshotNode')):

                    # If a previous drag of a Clip Creation Data Object has been cleared, the clipData.transcriptNum
                    # will be "0", which indicated that the current Drag is NOT a Clip Creation Data Object, 
                    # and therefore should be processed elsewhere.  If it is NOT "0", we should process it here.
                    if clipData.transcriptNum != 0:
                        CreateClip(clipData, self.dropData, self.tree, self.dropNode)
                        # Once the drop is done or rejected, we must clear the Clip Creation data out of the DropTarget.
                        # If we don't, this data will still be there if a Tree Node drag occurs, and there is no way in that
                        # circumstance to know which of the dragged objects to process!  Clearing avoids that problem.
                        self.ClearClipData()

                # Dropping Document Text onto a Collection, Quote, Clip, or Snapshot creates a Regular Quote.
                # See if the Drop Target is the correct Node Type.  The type comparison was added to get this working on the Mac.
                if (isinstance(quoteData, QuoteDragDropData)) and \
                   ((self.dropData.nodetype == 'CollectionNode') or \
                    (self.dropData.nodetype == 'QuoteNode') or \
                    (self.dropData.nodetype == 'ClipNode') or \
                    (self.dropData.nodetype == 'SnapshotNode')):

                    # If a previous drag of a Quote Creation Data Object has been cleared, the quoteData.documentNum
                    # will be "0", which indicated that the current Drag is NOT a Quote Creation Data Object, 
                    # and therefore should be processed elsewhere.  If it is NOT "0", we should process it here.
                    if quoteData.documentNum != 0:
                        CreateQuote(quoteData, self.dropData, self.tree, self.dropNode)
                        # Once the drop is done or rejected, we must clear the Quote Creation data out of the DropTarget.
                        # If we don't, this data will still be there if a Tree Node drag occurs, and there is no way in that
                        # circumstance to know which of the dragged objects to process!  Clearing avoids that problem.
                        self.ClearClipData()

                # Dropping Transcript Text onto a Keyword Group creates a Keyword.
                elif (type(clipData) == type(ClipDragDropData())) and \
                     (clipData.plainText != '') and \
                     (self.dropData.nodetype == 'KeywordGroupNode'):
                    # Create a new Keyword Object with the desired KWG and KW values
                    kw = Keyword.Keyword()
                    kw.keywordGroup = self.tree.GetItemText(self.dropNode)
                    # While the Clipboard's Plain Text has TIME CODES in it ...
                    while (clipData.plainText.find(u'\xa4') > -1) and \
                          (clipData.plainText.find('>', clipData.plainText.find(u'\xa4')) > 0):
                        # ... remove the time codes and the time code data
                        clipData.plainText = clipData.plainText[:clipData.plainText.find(u'\xa4')] + \
                                             clipData.plainText[clipData.plainText.find('>', clipData.plainText.find(u'\xa4')) + 1:]
                    # If there's still a time code, the data must have been truncated before the ">" terminator.
                    if (clipData.plainText.find(u'\xa4') > -1):
                        # Remove it from the end of the string.
                        clipData.plainText = clipData.plainText[:clipData.plainText.find(u'\xa4')]
                    # Limit the keyword length to 85 characters!
                    kw.keyword = clipData.plainText[:85]
                    # Create the Keyword Properties Dialog Box to Add a Keyword
                    dlg = KeywordPropertiesForm.EditKeywordDialog(None, -1, kw)
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
                                kw.db_save()
                                # Add the new Keyword to the tree
                                self.tree.add_Node('KeywordNode', (_('Keywords'), kw.keywordGroup, kw.keyword), 0, kw.keywordGroup)

                                # Now let's communicate with other Transana instances if we're in Multi-user mode
                                if not TransanaConstants.singleUserVersion:
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (kw.keywordGroup, kw.keyword))

                                # If we do all this, we don't need to continue any more.
                                contin = False
                            # If the user pressed Cancel ...
                            else:
                                # ... then we don't need to continue any more.
                                contin = False
                        # Handle "SaveError" exception
                        except TransanaExceptions.SaveError, e:
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
                    # Destroy the Keyword Dialog
                    dlg.Destroy()
                    # Clear the Clip Data
                    self.ClearClipData()

                # Dropping Document Text onto a Keyword Group creates a Keyword.
                elif (type(quoteData) == type(QuoteDragDropData())) and \
                     (quoteData.plainText != '') and \
                     (self.dropData.nodetype == 'KeywordGroupNode'):

                    print "DragAndDropObjects.DataTreeDropTarget.OnDrop():  Drop Document Text onto Keyword Group should create a keyword!"
                    print quoteData
                    print

                    # Create a new Keyword Object with the desired KWG and KW values
                    kw = Keyword.Keyword()
                    kw.keywordGroup = self.tree.GetItemText(self.dropNode)
                    # While the Clipboard's Plain Text has TIME CODES in it ...
                    # (If the Document was imported from an exported Transcript, it *could* have time codes!!
                    while (quoteData.plainText.find(u'\xa4') > -1) and \
                          (quoteData.plainText.find('>', quoteData.plainText.find(u'\xa4')) > 0):
                        # ... remove the time codes and the time code data
                        quoteData.plainText = quoteData.plainText[:quoteData.plainText.find(u'\xa4')] + \
                                              quoteData.plainText[quoeData.plainText.find('>', quoteData.plainText.find(u'\xa4')) + 1:]
                    # If there's still a time code, the data must have been truncated before the ">" terminator.
                    if (quoteData.plainText.find(u'\xa4') > -1):
                        # Remove it from the end of the string.
                        quoteData.plainText = quoteData.plainText[:quoteData.plainText.find(u'\xa4')]
                    # Limit the keyword length to 85 characters!
                    kw.keyword = quoteData.plainText[:85]
                    # Create the Keyword Properties Dialog Box to Add a Keyword
                    dlg = KeywordPropertiesForm.EditKeywordDialog(None, -1, kw)
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
                                kw.db_save()
                                # Add the new Keyword to the tree
                                self.tree.add_Node('KeywordNode', (_('Keywords'), kw.keywordGroup, kw.keyword), 0, kw.keywordGroup)

                                # Now let's communicate with other Transana instances if we're in Multi-user mode
                                if not TransanaConstants.singleUserVersion:
                                    if TransanaGlobal.chatWindow != None:
                                        TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (kw.keywordGroup, kw.keyword))

                                # If we do all this, we don't need to continue any more.
                                contin = False
                            # If the user pressed Cancel ...
                            else:
                                # ... then we don't need to continue any more.
                                contin = False
                        # Handle "SaveError" exception
                        except TransanaExceptions.SaveError, e:
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
                    # Destroy the Keyword Dialog
                    dlg.Destroy()
                    # Clear the Quote Data
                    self.ClearQuoteData()

                # Dropping Transcript Text onto a Keyword creates a Quick Clip.
                # If a previous drag of a Clip Creation Data Object has been cleared, the clipData.transcriptNum
                # will be "0", which indicated that the current Drag is NOT a Clip Creation Data Object, 
                # and therefore should be processed elsewhere.  If it is NOT "0", we should process it here.
                elif (type(clipData) == type(ClipDragDropData())) and \
                     (clipData.transcriptNum != 0) and \
                     (self.dropData.nodetype == 'KeywordNode'):
                    # Pass the accumulated data to the CreateQuickClip method, which is in the DragAndDropObjects module
                    # because drag and drop is an alternate way to create a Quick Clip.
                    CreateQuickClip(clipData, self.dropData.parent, self.tree.GetItemText(self.dropNode), self.tree)
                    # Once the drop is done, we must clear the Clip Creation data out of the DropTarget.
                    # If we don't, this data will still be there if a Tree Node drag occurs, and there is no way in that
                    # circumstance to know which of the dragged objects to process!  Clearing avoids that problem.
                    self.ClearClipData()

                # Dropping Document Text onto a Keyword creates a Quick Quote.
                # If a previous drag of a Quote Creation Data Object has been cleared, the QuoteData.documentNum
                # will be "0", which indicated that the current Drag is NOT a QUote Creation Data Object, 
                # and therefore should be processed elsewhere.  If it is NOT "0", we should process it here.
                elif (type(quoteData) == type(QuoteDragDropData())) and \
                     (quoteData.documentNum != 0) and \
                     (self.dropData.nodetype == 'KeywordNode'):
                    # Pass the accumulated data to the CreateQuickQuote method, which is in the DragAndDropObjects module
                    # because drag and drop is an alternate way to create a Quick Quote.
                    CreateQuickQuote(quoteData, self.dropData.parent, self.tree.GetItemText(self.dropNode), self.tree)
                    # Once the drop is done, we must clear the Clip Creation data out of the DropTarget.
                    # If we don't, this data will still be there if a Tree Node drag occurs, and there is no way in that
                    # circumstance to know which of the dragged objects to process!  Clearing avoids that problem.
                    self.ClearQuoteData()

                else:
                    # If the Drop target is not valid, we prevent the drop process by altering the wxDropResult (dragResult)
                    dragResult = wx.DragNone
                # Reset the cursor, regardless of whether the drop succeeded or failed.
                self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                 
            except:
                # Reset the cursor, regardless of whether the drop succeeded or failed.
                self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

                (exType, exValue) =  sys.exc_info()[:2]
         
                # If an expection occurs here, it's no big deal.  Forget about it.
                pass
            
        # Returning this value allows us to confirm or veto the drop request
        return dragResult


class ClipDragDropData(object):
    """ This object contains all the data that needs to be transferred in order to create a Clip
        from a selection in a Transcript. """

    def __init__(self, transcriptNum=0, episodeNum=0, clipStart=0, clipStop=0, text='', plainText='', videoCheckboxData=[]):
        """ ClipDragDropData Objects require the following parameters:
            transcriptNum      The Transcript Number of the originating Transcript
            episodeNum         The Episode the originating Transcript is attached to
            clipStart          The starting Time Code for the Clip
            clipStop           the ending Time Code for the Clip
            text               the Text for the Clip, in XML format
            videoCheckboxData  the Video Checkbox information from the Video Window. """
        self.transcriptNum = transcriptNum
        self.episodeNum = episodeNum
        self.clipStart = clipStart
        self.clipStop = clipStop
        self.text = text
        self.plainText = plainText
        self.videoCheckboxData = videoCheckboxData

    def __repr__(self):
        str = 'ClipDragDropData Object:\n'
        str = str + 'transcriptNum = %s\n' % self.transcriptNum
        str = str + 'episodeNum = %s\n' % self.episodeNum
        str = str + 'clipStart = %s\n' % Misc.time_in_ms_to_str(self.clipStart)
        str = str + 'clipStop = %s\n' % Misc.time_in_ms_to_str(self.clipStop)
        str = str + 'text = %s\n' % self.text
        str += 'plainText = %s\n\n' % self.plainText
        str += 'videoCheckboxData = %s\n\n' % self.videoCheckboxData
        return str

class QuoteDragDropData(object):
    """ This object contains all the data that needs to be transferred in order to create a Quote
        from a selection in a Document. """

    def __init__(self, documentNum=0, sourceQuote=0, startChar=0, endChar=0, text='', plainText=''):
        """ QuoteDragDropData Objects require the following parameters:
              documentNum      The Document Number of the originating Document
              sourceQuote      The Number of the Quote this is taken from, if it is from a Quote
              startChar        The starting character for the quote
              endChar          The ending character for the Quote
              text             the Text for the Quote, in XML format. """
        self.documentNum = documentNum
        self.sourceQuote = sourceQuote
        self.startChar = startChar
        self.endChar = endChar
        self.text = text
        self.plainText = plainText

    def __repr__(self):
        str = 'QuoteDragDropData Object:\n'
        str += 'documentNum = %s\n' % self.documentNum
        str += 'sourceQuote = %s\n' % self.sourceQuote
        str += 'startChar = %s\n' % self.startChar
        str += 'endChar = %s\n' % self.endChar
        str += 'text = %s\n' % self.text
        str += 'plainText = %s\n\n' % self.plainText
        return str


def CreateClip(clipData, dropData, tree, dropNode):
    """ This method handles the creation of a Clip Object in the Transana Database """
    # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
    contin = True
    # Create a new Clip Object
    tempClip = Clip.Clip()
    # We need to know if the Clip is coming from an Episode or another Clip.
    # We can determine that by looking at the transcript passed in the ClipData
    # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
    tempTranscript = Transcript.Transcript(clipData.transcriptNum, skipText=True)
    # If we are working from an Episode Transcript ...
    if tempTranscript.clip_num == 0:
        # Get the Episode Number from the clipData Object
        tempClip.episode_num = clipData.episodeNum

        # Get the Transcript Number from the clipData Object
        trNum = clipData.transcriptNum
    # If we are working from a Clip Transcript ...
    else:
        sourceClip = Clip.Clip(tempTranscript.clip_num)

        # Get the Episode Number from the sourceClip Object
        tempClip.episode_num = sourceClip.episode_num
        # Get the source transcript number from the clip transcript
        trNum = sourceClip.transcripts[0].source_transcript
    # Get the Clip Start Time from the clipData Object
    tempClip.clip_start = clipData.clipStart

    # Check to see if the clip starts before the media file starts (due to Adjust Indexes)
    if tempClip.clip_start < 0.0:
        prompt = _('The starting point for a Clip cannot be before the start of the media file.')
        errordlg = Dialogs.ErrorDialog(None, prompt)
        errordlg.ShowModal()
        errordlg.Destroy()
        # If so, cancel the clip creation
        return

    # Check to see if the clip starts after the media file ends (due to Adjust Indexes)
    if tempClip.clip_start >= TransanaGlobal.menuWindow.ControlObject.VideoWindow.GetMediaLength():
        prompt = _('The starting point for a Clip cannot be after the end of the media file.')
        errordlg = Dialogs.ErrorDialog(None, prompt)
        errordlg.ShowModal()
        errordlg.Destroy()
        # If so, cancel the Clip creation
        return

    # Get the Clip Stop Time from the clipData Object
    tempClip.clip_stop = clipData.clipStop

    # Check to see if the clip goes all the way to the end of the media file.  If so, it's probably an accident
    # due to the lack of an ending time code.  We'll skip this in the last 30 seconds of the media file, though.
    if (tempClip.clip_stop == TransanaGlobal.menuWindow.ControlObject.VideoWindow.GetMediaLength()) and \
       (tempClip.clip_stop - tempClip.clip_start > 30000):
        prompt = _('The ending point for this Clip is the end of the media file.  Do you want to create this clip?')
        errordlg = Dialogs.QuestionDialog(None, prompt, _("Transana Error"))
        result = errordlg.LocalShowModal()
        errordlg.Destroy()
        # If the user says NO, they don't want to create it ...
        if result == wx.ID_NO:
            # .. cancel Clip Creation
            return

    # Check to see if the clip ends after the media file ends (due to Adjust Indexes)
    if tempClip.clip_stop > TransanaGlobal.menuWindow.ControlObject.VideoWindow.GetMediaLength():
        prompt = _('The ending point for this Clip is after the end of the media file.  This clip may not end where you expect.')
        errordlg = Dialogs.ErrorDialog(None, prompt)
        errordlg.ShowModal()
        errordlg.Destroy()
        # We don't cancel clip creation, but we do adjust the end of the clip.
        tempClip.clip_stop = TransanaGlobal.menuWindow.ControlObject.VideoWindow.GetMediaLength()

    # Create a Transcript object
    tempClipTranscript = Transcript.Transcript()
    # Get the Episode Number
    tempClipTranscript.episode_num = tempClip.episode_num
    # Get the Source Transcript number
    tempClipTranscript.source_transcript = trNum
    # Get the Start Time
    tempClipTranscript.clip_start = clipData.clipStart
    # Get the Clip Stop Time
    tempClipTranscript.clip_stop = clipData.clipStop
    # Assign the Transcript Text
    tempClipTranscript.text = clipData.text
    # Add the Temporary Transcript to the Quick Clip
    tempClip.transcripts.append(tempClipTranscript)

    # If the Clip Creation Object is dropped on a Collection ...
    if dropData.nodetype == 'CollectionNode':
        # ... get the Clip's Collection Number from the Drop Node ...
        tempClip.collection_num = dropData.recNum
        # ... and the Clip's Collection Name from the Drop Node.
        tempClip.collection_id = tree.GetItemText(dropNode)
        # If dropping on a Collection, set Sort Order to the end of the list
        tempClip.sort_order = DBInterface.getMaxSortOrder(dropData.recNum) + 1        
        # Remember the Collection Node which should be the parent for the new Clip Node to be created later.
        collectionNode = dropNode
    # If the Clip Creation Object is dropped on a Quote, Clip or Snapshot ...
    elif dropData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
        # ... get the Clip's Collection Number from the Drop Node's Parent ...
        tempClip.collection_num = dropData.parent
        # ... and the Clip's Collection Name from the Drop Node's Parent.
        tempClip.collection_id = tree.GetItemText(tree.GetItemParent(dropNode))
        # Remember the Collection Node which should be the parent for the new Clip Node to be created later.
        collectionNode = tree.GetItemParent(dropNode)

    # Load the Episode that is connected to the Clip's Originating Transcript
    tempEpisode = Episode.Episode(tempClip.episode_num)
    # Start the clip off with the Episode's offset, though this could change if the first video wasn't used!
    tempClip.offset = tempEpisode.offset
    # Initially, assume that we don't need to shift the Clip offset, i.e. that the offset shift is ZERO
    offsetShift = 0
    # If there is no videoCheckbox Data (ie there are no video checkboxes) or the FIRST media files should be included ...
    if (clipData.videoCheckboxData == []) or (clipData.videoCheckboxData[0][0]):
        # The Clip's Media Filename comes from the Episode Record
        tempClip.media_filename = tempEpisode.media_filename
    # audio defaults to 1 (on).  If there are checkboxes and the first audio indicator is unchecked ...
    if (clipData.videoCheckboxData != []) and (not clipData.videoCheckboxData[0][1]):
        # ... then indicate that the first audio track should not be included.
        tempClip.audio = 0

    # For each set of media player checkboxes after the first (which has already been processed) ...
    for x in range(1, len(clipData.videoCheckboxData)):
        # ... get the checkbox data
        (videoCheck, audioCheck) = clipData.videoCheckboxData[x]
        # if the media should be included ...
        if videoCheck:
            # if this is the FIRST included media file, store the data in the Clip object.
            if tempClip.media_filename == '':
                # Grab the file name
                tempClip.media_filename = tempEpisode.additional_media_files[x - 1]['filename']
                # If we wind up here, we need to shift the offset values.  Remember the amount to shift them.
                offsetShift = tempEpisode.additional_media_files[x - 1]['offset']
                # Add the offset shift value to the Clip's gobal offset
                tempClip.offset += offsetShift
                # Note whether the audio should be played by default
                tempClip.audio = audioCheck
            # If this is NOT the first included media file, store the data in the additional_media_files structure,
            # adjusting the offset by the offsetShift value if needed.  (YES, minus here, plus above.)
            else:
                tempClip.additional_media_files = {'filename' : tempEpisode.additional_media_files[x - 1]['filename'],
                                                   'length'   : tempClip.clip_stop - tempClip.clip_start,
                                                   'offset'   : tempEpisode.additional_media_files[x - 1]['offset'] - offsetShift,
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

    # We need to set up the initial keywords.
    # If we are working from an Episode Transcript ...
    if tempTranscript.clip_num == 0:
        # Iterate through the source Episode's Keyword List ...
        for clipKeyword in tempEpisode.keyword_list:
            # ... and copy each keyword to the new Clip
            tempClip.add_keyword(clipKeyword.keywordGroup, clipKeyword.keyword)
    # If we are working from a Clip Transcript ...
    else:
        # Iterate through the source Clip's Keyword List ...
        for clipKeyword in sourceClip.keyword_list:
            # ... and copy each keyword to the new Clip
            tempClip.add_keyword(clipKeyword.keywordGroup, clipKeyword.keyword)

    # Load the parent Collection
    tempCollection = Collection.Collection(tempClip.collection_num)
    try:
        # Lock the parent Collection, to prevent it from being deleted out from under the add.
        tempCollection.lock_record()
        collectionLocked = True
    # Handle the exception if the record is already locked by someone else
    except TransanaExceptions.RecordLockedError, c:
        # If we can't get a lock on the Collection, it's really not that big a deal.  We only try to get it
        # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
        # still be able to add Clips even if someone else is editing the Collection properties.
        collectionLocked = False

    # Create the Clip Properties Dialog Box to Add a Clip
    dlg = ClipPropertiesForm.AddClipDialog(None, -1, tempClip)
    # While the "continue" flag is True ...
    while contin:
        # Display the Clip Properties Dialog Box and get the data from the user
        if dlg.get_input() != None:
            # Use "try", as exceptions could occur
            try:
                # See if the Clip Name already exists in the Destination Collection
                (dupResult, newClipName) = CheckForDuplicateObjName(tempClip.id, 'ClipNode', tree, dropNode)

                # If a Duplicate Clip Name is found and the error situation not resolved, show an Error Message
                if dupResult:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Clip Creation cancelled for Clip "%s".  Duplicate Item Name Error.'), 'utf8') % tempClip.id
                    else:
                        prompt = _('Clip Creation cancelled for Clip "%s".  Duplicate Item Name Error.') % tempClip.id
                    errordlg = Dialogs.ErrorDialog(None, prompt)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Unlock the parent collection
                    if collectionLocked:
                        tempCollection.unlock_record()
                    # If the user cancels Clip Creation, we don't need to continue any more.
                    contin = False
                else:
                    # Create a Popup Dialog.  (After duplicate names have been resolved to avoid conflict.)
                    tmpDlg = Dialogs.PopupDialog(None, _("Saving Clip"), _("Saving the Clip"))
                    # If the Name was changed, reflect that in the Clip Object
                    tempClip.id = newClipName
                    tempCollection = Collection.Collection(tempClip.collection_num)

                    # Try to save the data from the form
                    tempClip.db_save()
                    nodeData = (_('Collections'),) + tempCollection.GetNodeData() + (tempClip.id,)

                    # See if we're dropping on a Clip Node ...
                    if dropData.nodetype == 'ClipNode':
                        # Add the new Collection to the data tree
                        newNode = tree.add_Node('ClipNode', nodeData, tempClip.number, tempClip.collection_num, sortOrder=tempClip.sort_order, expandNode=True, insertPos=dropNode)
                    else:
                        # Add the new Clip to the data tree
                        tree.add_Node('ClipNode', nodeData, tempClip.number, tempClip.collection_num, sortOrder=tempClip.sort_order, avoidRecursiveYields=True)

                    # See if we're dropping on a Quote, Clip, or Snapshot Node ...
                    if dropData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
                        # ... and if so, change the Sort Order of the clips
                        if not ChangeClipOrder(tree, dropNode, tempClip, tempCollection):
                            pass

#                            tempClip.lock_record()
                            # If ChangeClipOrder fails, make sure the clip is at the end of the Clip List
#                            tempClip.sort_order = DBInterface.getMaxSortOrder(dropData.parent) + 1
                            # Try to save the data from the form
#                            tempClip.db_save()
#                            tempClip.unlock_record()
                            # Reset the order of objects in the node, no need to send MU messages
#                            tree.UpdateCollectionSortOrder(tree.GetItemParent(dropNode), sendMessage=False)

                        # When we dropped Transcript text on a Clip, the screen wouldn't update until we touched the Mouse!
                        # This fixes that.
                        try:
                            wx.YieldIfNeeded()
                        except:
                            pass

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "ACl %s"
                        data = (nodeData[1],)

                        for nd in nodeData[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                    # See if the Keyword visualization needs to be updated.
                    tree.parent.ControlObject.UpdateKeywordVisualization()
                    # Even if this computer doesn't need to update the keyword visualization others, might need to.
                    if not TransanaConstants.singleUserVersion:
                        # We need to update the Episode Keyword Visualization
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', tempEpisode.number, 0))
                    
                    # Unlock the parent collection
                    if collectionLocked:
                        tempCollection.unlock_record()
                    # Remove the Popup Dialog
                    tmpDlg.Destroy()
                    # If we do all this, we don't need to continue any more.
                    contin = False
            # Handle "SaveError" exception
            except TransanaExceptions.SaveError:
                # Remove the Popup Dialog
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
                # Remove the Popup Dialog
                tmpDlg.Destroy()
                # Display the Exception Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, "%s" % (sys.exc_info()[:2], ))
                errordlg.ShowModal()
                errordlg.Destroy()

                import traceback
                traceback.print_exc(file=sys.stdout)
                
        # If the user pressed Cancel ...
        else:
            # Unlock the parent collection
            if collectionLocked:
                tempCollection.unlock_record()
            # ... then we don't need to continue any more.
            contin = False
    dlg.Destroy()

def CreateQuote(quoteData, dropData, tree, dropNode):
    """ This method handles the creation of a Quote Object in the Transana Database """
    # Load the Source Document, regardless of whether we're quoting a Document or a Quote
    sourceDoc = Document.Document(quoteData.documentNum)
    # Start Exception Handling
    try:
        # Lock the Source Document, as we cannot Quote from a Document someone else is editing.
        sourceDoc.lock_record()
        # Note that the Source Document Lock was obtained in this method
        sourceDocLocked = True
    # Handle Record Lock problems
    except TransanaExceptions.RecordLockedError, e:
        # See if the Source Document is open by THIS user.  DO NOT EDIT THIS OBJECT.
        tmpObj = tree.parent.ControlObject.GetOpenDocumentObject(Document.Document, quoteData.documentNum)
        # If THIS user has the object open AND has the Record Lock ...
        if (tmpObj != None) and tmpObj.isLocked:
            sourceDocLocked = False
            # Save the Source Document / Quote.  This prevents user from abandoning changes that Quotes depend on!!
            tree.parent.ControlObject.SaveTranscript(prompt=False)
        else:
            # ... handle the Record Lock Exception
            TransanaExceptions.ReportRecordLockedException(_('Document'), sourceDoc.id, e)
            # ... and abandon Quote creation
            return

    # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
    contin = True
    # Create a new Quote Object
    tempQuote = Quote.Quote()
    # Get the Quote's source document number form the quoteData
    tempQuote.source_document_num = quoteData.documentNum
    # Get the Quote Start Character from the quoteData Object
    tempQuote.start_char = quoteData.startChar
    # Get the Quote's End Character from the quoteData Object
    tempQuote.end_char = quoteData.endChar
    # Get the Quote's text from the quoteData Object
    tempQuote.text = quoteData.text

    # If the Quote Creation Object is dropped on a Collection ...
    if dropData.nodetype == 'CollectionNode':
        # ... get the Quote's Collection Number from the Drop Node ...
        tempQuote.collection_num = dropData.recNum
        # If dropping on a Collection, set Sort Order to the end of the list
        tempQuote.sort_order = DBInterface.getMaxSortOrder(dropData.recNum) + 1        
        # Remember the Collection Node which should be the parent for the new Quote Node to be created later.
        collectionNode = dropNode
    # If the Quote Creation Object is dropped on a Quote, Clip or Snapshot ...
    elif dropData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
        # ... get the Quote's Collection Number from the Drop Node's Parent ...
        tempQuote.collection_num = dropData.parent
        # Remember the Collection Node which should be the parent for the new Quote Node to be created later.
        collectionNode = tree.GetItemParent(dropNode)

    # We need to set up the initial keywords.
    # If we are working from a Document ...
    if quoteData.sourceQuote == 0:
        # Load the Source Document
        tmpSource = sourceDoc
    else:
        tmpSource = Quote.Quote(num=quoteData.sourceQuote)
    # Iterate through the source Document's Keyword List ...
    for quoteKeyword in tmpSource.keyword_list:
        # ... and copy each keyword to the new Quote
        tempQuote.add_keyword(quoteKeyword.keywordGroup, quoteKeyword.keyword)

    # Load the parent Collection
    tempCollection = Collection.Collection(tempQuote.collection_num)
    try:
        # Lock the parent Collection, to prevent it from being deleted out from under the add.
        tempCollection.lock_record()
        collectionLocked = True
    # Handle the exception if the record is already locked by someone else
    except TransanaExceptions.RecordLockedError, c:
        # If we can't get a lock on the Collection, it's really not that big a deal.  We only try to get it
        # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
        # still be able to add Quotes even if someone else is editing the Collection properties.
        collectionLocked = False

    # Create the Quote Properties Dialog Box to Add a Quote
    dlg = QuotePropertiesForm.AddQuoteDialog(None, -1, tempQuote)
    # While the "continue" flag is True ...
    while contin:
        # Display the Quote Properties Dialog Box and get the data from the user
        if dlg.get_input() != None:
            # Use "try", as exceptions could occur
            try:
                # See if the Quote Name already exists in the Destination Collection
                (dupResult, newQuoteName) = CheckForDuplicateObjName(tempQuote.id, 'QuoteNode', tree, dropNode)

                # If a Duplicate Clip Name is found and the error situation not resolved, show an Error Message
                if dupResult:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Quote Creation cancelled for Quote "%s".  Duplicate Item Name Error.'), 'utf8') % tempQuote.id
                    else:
                        prompt = _('Quote Creation cancelled for Quote "%s".  Duplicate Item Name Error.') % tempQuote.id
                    errordlg = Dialogs.ErrorDialog(None, prompt)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Unlock the parent collection
                    if collectionLocked:
                        tempCollection.unlock_record()
                    # If the user cancels Clip Creation, we don't need to continue any more.
                    contin = False
                else:
                    # Create a Popup Dialog.  (After duplicate names have been resolved to avoid conflict.)
                    tmpDlg = Dialogs.PopupDialog(None, _("Saving Quote"), _("Saving the Quote"))
                    # If the Name was changed, reflect that in the Quote Object
                    tempQuote.id = newQuoteName
                    tempCollection = Collection.Collection(tempQuote.collection_num)

                    # Try to save the data from the form
                    tempQuote.db_save()
                    # Inform the Control Object of the new Quote so it can update the open Document!
                    tree.parent.ControlObject.AddQuoteToOpenDocument(tempQuote)
                        
                    nodeData = (_('Collections'),) + tempCollection.GetNodeData() + (tempQuote.id,)

                    # See if we're dropping on a Quote Node ...
                    if dropData.nodetype == 'QuoteNode':
                        # Add the new Collection to the data tree
                        newNode = tree.add_Node('QuoteNode', nodeData, tempQuote.number, tempQuote.collection_num, sortOrder=tempQuote.sort_order, expandNode=True, insertPos=dropNode)
                    else:
                        # Add the new Quote to the data tree
                        tree.add_Node('QuoteNode', nodeData, tempQuote.number, tempQuote.collection_num, sortOrder=tempQuote.sort_order, avoidRecursiveYields=True)

                    # See if we're dropping on a Quote, Clip, or Snapshot Node ...
                    if dropData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
                        # ... and if so, change the Sort Order of the items
                        if not ChangeClipOrder(tree, dropNode, tempQuote, tempCollection):
                            pass

                        # Sort the PARENT (Collection) node
                        tree.SortChildren(tree.GetItemParent(dropNode))

                        # When we dropped Transcript text on a Quote, Clips, or Snapsho, the screen wouldn't update until
                        # we touched the Mouse!
                        # This fixes that.
                        try:
                            wx.YieldIfNeeded()
                        except:
                            pass

                    else:
                        # Sort the DROP (Collection) node
                        tree.SortChildren(dropNode)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "AQ %s"
                        data = (nodeData[1],)

                        for nd in nodeData[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                    # See if the Keyword visualization needs to be updated.
                    tree.parent.ControlObject.UpdateKeywordVisualization()
                    # Even if this computer doesn't need to update the keyword visualization others, might need to.
                    if not TransanaConstants.singleUserVersion:
                        # We need to update the Document Keyword Visualization
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Document', 0, tempQuote.source_document_num))
                    
                    # Unlock the parent collection
                    if collectionLocked:
                        tempCollection.unlock_record()
                    # Remove the Popup Dialog
                    tmpDlg.Destroy()
                    # If we do all this, we don't need to continue any more.
                    contin = False
            # Handle "SaveError" exception
            except TransanaExceptions.SaveError:
                # Remove the Popup Dialog
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
                # Remove the Popup Dialog
                tmpDlg.Destroy()
                # Display the Exception Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, "%s" % (sys.exc_info()[:2], ))
                errordlg.ShowModal()
                errordlg.Destroy()

                import traceback
                traceback.print_exc(file=sys.stdout)
                
        # If the user pressed Cancel ...
        else:
            # Unlock the parent collection
            if collectionLocked:
                tempCollection.unlock_record()
            # ... then we don't need to continue any more.
            contin = False
    dlg.Destroy()
    # If the Source Doc was locked in this method ...
    if sourceDocLocked:
        # ... then unlock the Source Document
        sourceDoc.unlock_record()

def DropKeyword(parent, sourceData, targetType, targetName, targetRecNum, targetParent, confirmations=True):
    """Drop a Keyword onto an Object.  sourceData is from the Keyword.  The
    targetType is one of 'Libraries', 'Document', 'Episode', 'Collection', or 'Clip'.
    targetParent is only used for collections."""

    if targetType == 'Libraries':
        # Prompt for confirmation if that is desired
        if confirmations:
            # Get user confirmation of the Keyword Add request
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to add Keyword "%s:%s" to all the Items in\nLibrary "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
            dlg = Dialogs.QuestionDialog(parent, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            if result == wx.ID_NO:
                return
        # If confirmed, copy the Keyword to all Episodes in the Library

        # First, let's load the Library Record
        tempLibrary = Library.Library(targetRecNum)
        try:
            # Lock the Library Record, just to be on the safe side (Is this necessary??  I don't think so, but maybe that can confirm that all episodes are available.)
            tempLibrary.lock_record()
            # Now get a list of all Documents in the Library and iterate through them
            for tempDocumentNum, tempDocumentID, tempLibraryNum in DBInterface.list_of_documents(tempLibrary.number):
                # Load the Document Record
                tempDocument = Document.Document(num=tempDocumentNum)
                try:
                    # Lock the Document Record
                    tempDocument.lock_record()
                    # Add the Keyword to the Episode
                    tempDocument.add_keyword(sourceData.parent, sourceData.text)
                    # If we're using confirmations (only have ONE operation) ...
                    if confirmations:
                        # ... Check to see if there are keywords to be propagated
                        parent.parent.ControlObject.PropagateObjectKeywords(_('Document'), tempDocument.number, tempDocument.keyword_list)
                    # Save the Document
                    tempDocument.db_save()
                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = 'Document %d' % tempDocument.number
                        if TransanaGlobal.chatWindow != None:
                            # Send the "Update Keyword List" message
                            TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
                    # Unlock the Document Record
                    tempDocument.unlock_record()
                # Handle "RecordLockedError" exception
                except TransanaExceptions.RecordLockedError, e:
                    TransanaExceptions.ReportRecordLockedException(_("Document"), tempDocument.id, e)
            # Now get a list of all Episodes in the Library and iterate through them
            for tempEpisodeNum, tempEpisodeID, tempLibraryNum in DBInterface.list_of_episodes_for_series(tempLibrary.id):
                # Load the Episode Record
                tempEpisode = Episode.Episode(num=tempEpisodeNum)
                try:
                    # Lock the Episode Record
                    tempEpisode.lock_record()
                    # Add the Keyword to the Episode
                    tempEpisode.add_keyword(sourceData.parent, sourceData.text)
                    # If we're using confirmations (only have ONE operation) ...
                    if confirmations:
                        # ... Check to see if there are keywords to be propagated
                        parent.parent.ControlObject.PropagateObjectKeywords(_('Episode'), tempEpisode.number, tempEpisode.keyword_list)
                    # Save the Episode
                    tempEpisode.db_save()
                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = 'Episode %d' % tempEpisode.number
                        if TransanaGlobal.chatWindow != None:
                            # Send the "Update Keyword List" message
                            TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
                    # Unlock the Episode Record
                    tempEpisode.unlock_record()
                # Handle "RecordLockedError" exception
                except TransanaExceptions.RecordLockedError, e:
                    TransanaExceptions.ReportRecordLockedException(_("Episode"), tempEpisode.id, e)
            # Unlock the Library Record
            tempLibrary.unlock_record()   
        # Handle "RecordLockedError" exception
        except TransanaExceptions.RecordLockedError, e:
            TransanaExceptions.ReportRecordLockedException(_('Libraries'), tempLibrary.id, e)
    
    elif targetType == 'Document':
        # Prompt for confirmation if that is desired
        if confirmations:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to add Keyword "%s:%s" to\nDocument "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
            else:
                prompt = _('Do you want to add Keyword "%s:%s" to\nDocument "%s"?') % (sourceData.parent, sourceData.text, targetName)
            dlg = Dialogs.QuestionDialog(parent, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            if result == wx.ID_NO:
                return
        # If confirmed, copy the Keyword to the Documents

        # Load the Document Record
        tempDocument = Document.Document(num=targetRecNum)
        try:
            # Lock the Document Record
            tempDocument.lock_record()
            # Add the keyword to the Document
            tempDocument.add_keyword(sourceData.parent, sourceData.text)
            # If we're using confirmations (only have ONE operation) ...
            if confirmations:
                # ... Check to see if there are keywords to be propagated
                parent.parent.ControlObject.PropagateObjectKeywords(_('Document'), tempDocument.number, tempDocument.keyword_list)
            # Save the Document
            tempDocument.db_save()
            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion:
                msg = 'Document %d' % tempDocument.number
                if TransanaGlobal.chatWindow != None:
                    # Send the "Update Keyword List" message
                    TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
            # Unlock the Document
            tempDocument.unlock_record()
        # Handle "RecordLockedError" exception
        except TransanaExceptions.RecordLockedError, e:
            TransanaExceptions.ReportRecordLockedException(_("Document"), tempDocument.id, e)
    
    elif targetType == 'Episode':
        # Prompt for confirmation if that is desired
        if confirmations:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to add Keyword "%s:%s" to\nEpisode "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
            else:
                prompt = _('Do you want to add Keyword "%s:%s" to\nEpisode "%s"?') % (sourceData.parent, sourceData.text, targetName)
            dlg = Dialogs.QuestionDialog(parent, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            if result == wx.ID_NO:
                return
        # If confirmed, copy the Keyword to the Episodes

        # Load the Episode Record
        tempEpisode = Episode.Episode(num=targetRecNum)
        try:
            # Lock the Episode Record
            tempEpisode.lock_record()
            # Add the keyword to the Episode
            tempEpisode.add_keyword(sourceData.parent, sourceData.text)
            # If we're using confirmations (only have ONE operation) ...
            if confirmations:
                # ... Check to see if there are keywords to be propagated
                parent.parent.ControlObject.PropagateObjectKeywords(_('Episode'), tempEpisode.number, tempEpisode.keyword_list)
            # Save the Episode
            tempEpisode.db_save()
            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion:
                msg = 'Episode %d' % tempEpisode.number
                if TransanaGlobal.chatWindow != None:
                    # Send the "Update Keyword List" message
                    TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
            # Unlock the Episode
            tempEpisode.unlock_record()
        # Handle "RecordLockedError" exception
        except TransanaExceptions.RecordLockedError, e:
            TransanaExceptions.ReportRecordLockedException(_("Episode"), tempEpisode.id, e)
    
    elif targetType == 'Collection':
        # Prompt for confirmation if that is desired
        if confirmations:
            # Get user confirmation of the Keyword Add request
            prompt = unicode(_('Do you want to add Keyword "%s:%s" to all Items in\nCollection "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
            dlg = Dialogs.QuestionDialog(parent, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            if result == wx.ID_NO:
                return
        # If confirmed, copy the Keyword to all Clips in the Collection

        # We need a flag indicating if we need to update the Keyword Visualization
        updateKeywordVisualization = False
        # First, load the Collection
        tempCollection = Collection.Collection(targetRecNum, targetParent)
        try:
            # Lock the Collection Record, just to be on the safe side (Is this necessary??  I don't think so, but maybe that can confirm that all Clips are available.)
            tempCollection.lock_record()

            if TransanaConstants.proVersion:
                # Now load a list of all the Quotes in the Collection and iterate through them
                for tempQuoteNum, tempQuoteID, tempCollectNum in DBInterface.list_of_quotes_by_collectionnum(tempCollection.number):
                    # Load the Quote.
                    tempQuote = Quote.Quote(num=tempQuoteNum)
                    try:
                        # Lock the Quote
                        tempQuote.lock_record()
                        # Add the Keyword to the Quote
                        tempQuote.add_keyword(sourceData.parent, sourceData.text)
                        # Save the Quote
                        tempQuote.db_save()
                        # If the affected Quote is for the current Document, we need to update the
                        # Keyword Visualization
                        if (isinstance(parent.parent.ControlObject.currentObj, Document.Document)) and \
                           (tempQuote.source_document_num == parent.parent.ControlObject.currentObj.number):
                            # Signal that the Keyword Visualization needs to be updated
                            updateKeywordVisualization = True
                        # If the affected Quote is the current Quote, we need to update the
                        # Keyword Visualization
                        elif (isinstance(parent.parent.ControlObject.currentObj, Quote.Quote)) and \
                           (tempQuote.number == parent.parent.ControlObject.currentObj.number):
                            # Signal that the Keyword Visualization needs to be updated
                            updateKeywordVisualization = True
                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            msg = 'Quote %d' % tempQuote.number
                            if TransanaGlobal.chatWindow != None:
                                # Send the "Update Keyword List" message
                                TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
                        # Unlock the Quote
                        tempQuote.unlock_record()
                    # Handle "RecordLockedError" exception
                    except TransanaExceptions.RecordLockedError, e:
                        TransanaExceptions.ReportRecordLockedException(_("Quote"), tempQuote.id, e)

            # Now load a list of all the Clips in the Collection and iterate through them
            for tempClipNum, tempClipID, tempCollectNum in DBInterface.list_of_clips_by_collection(tempCollection.id, tempCollection.parent):
                # Load the Clip.
                tempClip = Clip.Clip(id_or_num=tempClipNum)
                try:
                    # Lock the Clip
                    tempClip.lock_record()
                    # Add the Keyword to the Clip
                    tempClip.add_keyword(sourceData.parent, sourceData.text)
                    # Save the Clip
                    tempClip.db_save()
                    # If the affected clip is for the current Episode, we need to update the
                    # Keyword Visualization
                    if (isinstance(parent.parent.ControlObject.currentObj, Episode.Episode)) and \
                       (tempClip.episode_num == parent.parent.ControlObject.currentObj.number):
                        # Signal that the Keyword Visualization needs to be updated
                        updateKeywordVisualization = True
                    # If the affected clip is the current Clip, we need to update the
                    # Keyword Visualization
                    elif (isinstance(parent.parent.ControlObject.currentObj, Clip.Clip)) and \
                       (tempClip.number == parent.parent.ControlObject.currentObj.number):
                        # Signal that the Keyword Visualization needs to be updated
                        updateKeywordVisualization = True
                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = 'Clip %d' % tempClip.number
                        if TransanaGlobal.chatWindow != None:
                            # Send the "Update Keyword List" message
                            TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
                    # Unlock the Clip
                    tempClip.unlock_record()
                # Handle "RecordLockedError" exception
                except TransanaExceptions.RecordLockedError, e:
                    TransanaExceptions.ReportRecordLockedException(_("Clip"), tempClip.id, e)

            if TransanaConstants.proVersion:
                # Now load a list of all the Snapshots in the Collection and iterate through them
                for tempSnapshotNum, tempSnapshotID, tempCollectNum in DBInterface.list_of_snapshots_by_collectionnum(tempCollection.number):
                    # Load the Snapshot
                    tempSnapshot = Snapshot.Snapshot(tempSnapshotNum)
                    try:
                        # Lock the Snapshot
                        tempSnapshot.lock_record()
                        # Add the Keyword to the Snapshot
                        tempSnapshot.add_keyword(sourceData.parent, sourceData.text)
                        # Save the Snapshot
                        tempSnapshot.db_save()

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            msg = 'Snapshot %d' % tempSnapshot.number
                            if TransanaGlobal.chatWindow != None:
                                # Send the "Update Keyword List" message
                                TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
                        # Unlock the Snapshot
                        tempSnapshot.unlock_record()
                    # Handle "RecordLockedError" exception
                    except TransanaExceptions.RecordLockedError, e:
                        TransanaExceptions.ReportRecordLockedException(_("Snapshot"), tempSnapshot.id, e)

            # Unlock the Collection Record
            tempCollection.unlock_record()
            # If we need to update the Keyword Visualization, do so
            if updateKeywordVisualization:
                # See if the Keyword visualization needs to be updated.
                parent.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Keyword Visualization no matter what here, when adding a keyword to a Collection
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))
        # Handle "RecordLockedError" exception
        except TransanaExceptions.RecordLockedError, e:
            TransanaExceptions.ReportRecordLockedException(_("Collection"), tempCollection.id, e)

    elif targetType == 'Quote':
        # Prompt for confirmation if that is desired
        if confirmations:
            # Get user confirmation of the Keyword Add request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to add Keyword "%s:%s" to\nQuote "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
            else:
                prompt = _('Do you want to add Keyword "%s:%s" to\nQuote "%s"?') % (sourceData.parent, sourceData.text, targetName)
            dlg = Dialogs.QuestionDialog(parent, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            if result == wx.ID_NO:
                return
        try:
            # If confirmed, copy the Keyword to the Quote
            # First, load the Quote.
            tempQuote = Quote.Quote(num=targetRecNum)
            # Lock the Quote Record
            tempQuote.lock_record()
            # Add the Keyword to the Quote Record
            tempQuote.add_keyword(sourceData.parent, sourceData.text)

            # Save the Quote Record
            tempQuote.db_save()
            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion:
                msg = 'Quote %d' % tempQuote.number
                if TransanaGlobal.chatWindow != None:
                    # Send the "Update Keyword List" message
                    TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
            # Unlock the Clip Record
            tempQuote.unlock_record()
            # If the affected quote is for the current Document, we need to update the
            # Keyword Visualization
            if (isinstance(parent.parent.ControlObject.currentObj, Document.Document)) and \
               (tempQuote.source_document_num == parent.parent.ControlObject.currentObj.number):
                # See if the Keyword visualization needs to be updated.
                parent.parent.ControlObject.UpdateKeywordVisualization()
            # If the affected quote is the current Quote, we need to update the
            # Keyword Visualization
            if (isinstance(parent.parent.ControlObject.currentObj, Quote.Quote)) and \
               (tempQuote.number == parent.parent.ControlObject.currentObj.number):
                # See if the Keyword visualization needs to be updated.
                parent.parent.ControlObject.UpdateKeywordVisualization()
            # Even if this computer doesn't need to update the keyword visualization others, might need to.
            if not TransanaConstants.singleUserVersion:
                # We need to update the quote and document Keyword Visualization when adding a keyword to a Quote
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Quote', tempQuote.number, tempQuote.source_document_num))

        except TransanaExceptions.RecordLockedError, e:
            TransanaExceptions.ReportRecordLockedException(_("Quote"), tempQuote.id, e)
        # Handle "SaveError" exception
        except TransanaExceptions.SaveError:
            # Display the Error Message
            errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
            errordlg.ShowModal()
            errordlg.Destroy()
            # Unlock the Clip Record
            tempClip.unlock_record()
        # Handle other exceptions
        except:
            # Display the Exception Message
            errordlg = Dialogs.ErrorDialog(None, "%s" % (sys.exc_info()[:2]))
            errordlg.ShowModal()
            errordlg.Destroy()
            # Unlock the Clip Record
            tempClip.unlock_record()

    elif targetType == 'Clip':
        # Prompt for confirmation if that is desired
        if confirmations:
            # Get user confirmation of the Keyword Add request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to add Keyword "%s:%s" to\nClip "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
            else:
                prompt = _('Do you want to add Keyword "%s:%s" to\nClip "%s"?') % (sourceData.parent, sourceData.text, targetName)
            dlg = Dialogs.QuestionDialog(parent, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            if result == wx.ID_NO:
                return
        try:
            # If confirmed, copy the Keyword to the Clip
            # First, load the Clip.
            tempClip = Clip.Clip(id_or_num=targetRecNum)
            # Lock the Clip Record
            tempClip.lock_record()
            # Add the Keyword to the Clip Record
            tempClip.add_keyword(sourceData.parent, sourceData.text)

            # Save the Clip Record
            tempClip.db_save()
            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion:
                msg = 'Clip %d' % tempClip.number
                if TransanaGlobal.chatWindow != None:
                    # Send the "Update Keyword List" message
                    TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
            # Unlock the Clip Record
            tempClip.unlock_record()
            # If the affected clip is for the current Episode, we need to update the
            # Keyword Visualization
            if (isinstance(parent.parent.ControlObject.currentObj, Episode.Episode)) and \
               (tempClip.episode_num == parent.parent.ControlObject.currentObj.number):
                # See if the Keyword visualization needs to be updated.
                parent.parent.ControlObject.UpdateKeywordVisualization()
            # If the affected clip is the current Clip, we need to update the
            # Keyword Visualization
            if (isinstance(parent.parent.ControlObject.currentObj, Clip.Clip)) and \
               (tempClip.number == parent.parent.ControlObject.currentObj.number):
                # See if the Keyword visualization needs to be updated.
                parent.parent.ControlObject.UpdateKeywordVisualization()
            # Even if this computer doesn't need to update the keyword visualization others, might need to.
            if not TransanaConstants.singleUserVersion:
                # We need to update the Clip Keyword Visualization when adding a keyword to a clip
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', tempClip.number, tempClip.episode_num))

        except TransanaExceptions.RecordLockedError, e:
            TransanaExceptions.ReportRecordLockedException(_("Clip"), tempClip.id, e)
        # Handle "SaveError" exception
        except TransanaExceptions.SaveError:
            # Display the Error Message
            errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
            errordlg.ShowModal()
            errordlg.Destroy()
            # Unlock the Clip Record
            tempClip.unlock_record()
        # Handle other exceptions
        except:
            # Display the Exception Message
            errordlg = Dialogs.ErrorDialog(None, "%s" % (sys.exc_info()[:2]))
            errordlg.ShowModal()
            errordlg.Destroy()
            # Unlock the Clip Record
            tempClip.unlock_record()

    elif targetType == 'Snapshot':
        # Prompt for confirmation if that is desired
        if confirmations:
            # Get user confirmation of the Keyword Add request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to add Keyword "%s:%s" to\nSnapshot "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
            else:
                prompt = _('Do you want to add Keyword "%s:%s" to\nSnapshot "%s"?') % (sourceData.parent, sourceData.text, targetName)
            dlg = Dialogs.QuestionDialog(parent, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            if result == wx.ID_NO:
                return
        try:
            # If confirmed, copy the Keyword to the Snapshot
            # First, load the Snapshot.
            tempSnapshot = Snapshot.Snapshot(targetRecNum)
            # Lock the Snapshot Record
            tempSnapshot.lock_record()
            # Add the Keyword to the Snapshot Record
            tempSnapshot.add_keyword(sourceData.parent, sourceData.text)
            # Save the Snapshot Record
            tempSnapshot.db_save()

            # See if the Keyword visualization needs to be updated.
            parent.parent.ControlObject.UpdateKeywordVisualization()
            # Even if this computer doesn't need to update the keyword visualization others, might need to.
            if not TransanaConstants.singleUserVersion:
                # We need to update the Episode Keyword Visualization
                if (TransanaGlobal.chatWindow != None) and (tempSnapshot.episode_num != 0):
                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', tempSnapshot.episode_num, 0))
            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion:
                msg = 'Snapshot %d' % tempSnapshot.number
                if TransanaGlobal.chatWindow != None:
                    # Send the "Update Keyword List" message
                    TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)
            # Unlock the Snapshot Record
            tempSnapshot.unlock_record()
        except TransanaExceptions.RecordLockedError, e:
            TransanaExceptions.ReportRecordLockedException(_("Snapshot"), tempSnapshot.id, e)
        # Handle "SaveError" exception
        except TransanaExceptions.SaveError:
            # Display the Error Message
            errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
            errordlg.ShowModal()
            errordlg.Destroy()
            # Unlock the Snapshot Record
            tempSnapshot.unlock_record()
        # Handle other exceptions
        except:
            # Display the Exception Message
            errordlg = Dialogs.ErrorDialog(None, "%s" % (sys.exc_info()[:2]))
            errordlg.ShowModal()
            errordlg.Destroy()
            # Unlock the Snapshot Record
            tempSnapshot.unlock_record()

def ProcessPasteDrop(treeCtrl, sourceData, destNode, action, confirmations=True):
    """ This method processes a "Paste" or "Drop" request for the Transana Database Tree.
        Parameters are:
          treeCtrl   -- the wxTreeCtrl where the Paste or Drop is occurring (the DBTree)
          sourceData -- the DATA associated with the Cut/Copy or Drag, the _NodeData or the DataTreeDragDropData Object
          destNode   -- the actual Tree Node selected for Drop or Paste
          action     -- a string of "Copy" or "Move", indicating whether a Copy or Cut/Move has been requested.
                        (This value is ignored in some instances where "Move" has no meaning.  """

    global YESTOALL

    # Since we get the actual destination node as a parameter, let's first extract the Node Data for the Destination
    destNodeData = treeCtrl.GetPyData(destNode)

#    print "DragAndDropObjects.ProcessPasteDrop(): dropping %s on %s" % (sourceData.nodetype, destNodeData.nodetype)
    
    # Determine whether a Copy or Move is being requested, and set the appropriate Prompt Text
    if 'unicode' in wx.PlatformInfo:
        if action == 'Copy':
            copyMovePrompt = unicode(_('COPY'), 'utf8')
        elif action == 'Move':
            copyMovePrompt = unicode(_('MOVE'), 'utf8')
    else: 
        if action == 'Copy':
            copyMovePrompt = _('COPY')
        elif action == 'Move':
            copyMovePrompt = _('MOVE')


    # If the SOURCE data is a QUOTE ...
    if (sourceData.nodetype == 'QuoteNode'):
        # Start exception handling
        try:
            # See if the Quote exists (hasn't been deleted, which can happen after COPY)
            # Don't load the Quote text to save time.
            tmpQuote = Quote.Quote(sourceData.recNum, skipText=True)
        # If the Quote doesn't exist ...
        except:
            # ... then we can't paste it, can we?
            return

    # If the SOURCE data is a CLIP ...
    if (sourceData.nodetype == 'ClipNode'):
        # Start exception handling
        try:
            # See if the clip exists (hasn't been deleted, which can happen after COPY)
            # Don't load the Clip Transcript to save time.
            tmpClip = Clip.Clip(sourceData.recNum, skipText=True)
        # If the clip doesn't exist ...
        except:
            # ... then we can't paste it, can we?
            return

    # Drop a Document on a Library (Move a Document)
    if (sourceData.nodetype == 'DocumentNode' and destNodeData.nodetype == 'LibraryNode'):
        # Get the Document data
        tmpDocument = Document.Document(sourceData.recNum)
        # Get the data for the SOURCE Library
        oldLibrary = Library.Library(sourceData.parent)
        # Get the data for the Destination Library
        tmpLibrary = Library.Library(destNodeData.recNum)
        # Let's make sure the drop is on a DIFFERENT Library.  Otherwise, the Document record disappears from the tree until the
        # database is refreshed.
        if oldLibrary.number != tmpLibrary.number:
            # Begin Exception Handling for the Lock
            try:
                # Try to lock the Document
                tmpDocument.lock_record()
                # If successful, assign the DESTINATION Library Number to the Document.
                tmpDocument.library_num = tmpLibrary.number
                # Start nested Exception Handling for the Save
                try:
                    # Try to Save the Document
                    tmpDocument.db_save()

                    # Now we need to move the entry in the Data Tree
                    # Start by building the Document's new Node List
                    nodeData = (_('Libraries'), tmpLibrary.id, tmpDocument.id)
                    # Add the Document node to the data tree
                    treeCtrl.add_Node('DocumentNode', nodeData, tmpDocument.number, tmpLibrary.number, expandNode=False)
                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("AD %s >|< %s" % (nodeData[-2], nodeData[-1]))

                    # Now request that the Document's OLD node be deleted.
                    # First build the Node List for the OLD Document ...
                    nodeData = (_('Libraries'), oldLibrary.id, tmpDocument.id)
                    # ... and delete it from the tree
                    treeCtrl.delete_Node(nodeData, 'DocumentNode')

                    # If we are moving a Document, the Document's Notes need to travel with the Document.  The first step is to
                    # get a list of those Notes.
                    noteList = DBInterface.list_of_notes(Document=tmpDocument.number)
                    # If there are Document Notes, we need to make sure they travel with the Document
                    if noteList != []:
                        # Build the Node List for the new Document
                        nodeData = (_('Libraries'), tmpLibrary.id, tmpDocument.id)
                        # Select the new Document Node
                        newNode = treeCtrl.select_Node(nodeData, 'DocumentNode')
                        # Use the TreeCtrl's "add_note_nodes" method to move the notes locally
                        treeCtrl.add_note_nodes(noteList, newNode, Document=tmpDocument.number)
                        treeCtrl.Refresh()
                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            # Iterate through the Notes List
                            for noteid in noteList:
                                # Construct the message and data to be passed
                                msg = "ADN %s"
                                # To avoid problems in mixed-language environments, we need the UNTRANSLATED string here!
                                data = (u'Libraries',) + nodeData[1:]  + (noteid,)
                                # Build the message to be sent
                                for nd in data[1:]:
                                    msg += " >|< %s"
                                # Send the message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage(msg % data)

                # If the Save fails ...
                except TransanaExceptions.SaveError, e:
                     # Display the Error Message
                     msg = _('A Document named "%s" already exists in Library "%s".')
                     if 'unicode' in wx.PlatformInfo:
                         msg = unicode(msg, 'utf8')
                     errordlg = Dialogs.ErrorDialog(None, msg % (tmpDocument.id, Library.Library(destNodeData.recNum).id))
                     errordlg.ShowModal()
                     errordlg.Destroy()
                    
                # If we get this far, unlock the Document
                tmpDocument.unlock_record()
            # If we are unable to lock the Document ...
            except TransanaExceptions.RecordLockedError, e:
                # Report the Record Lock failure
                TransanaExceptions.ReportRecordLockedException(_('Document'), tmpDocument.id, e)
       
    # Drop an Episode on a Library (Move an Episode)
    elif (sourceData.nodetype == 'EpisodeNode' and destNodeData.nodetype == 'LibraryNode'):
        # Get the Episode data
        tmpEpisode = Episode.Episode(sourceData.recNum)
        # Get the data for the SOURCE Library
        oldLibrary = Library.Library(sourceData.parent)
        # Get the data for the Destination Library
        tmpLibrary = Library.Library(destNodeData.recNum)
        # Let's make sure the drop is on a DIFFERENT Library.  Otherwise, the Episode record disappears from the tree until the
        # database is refreshed.
        if oldLibrary.number != tmpLibrary.number:
            # Begin Exception Handling for the Lock
            try:
                # Try to lock the Episode
                tmpEpisode.lock_record()
                # If successful, assign the DESTINATION Library Number and ID to the Episode.
                tmpEpisode.series_num = tmpLibrary.number
                tmpEpisode.series_id = tmpLibrary.id
                # Start nested Exception Handling for the Save
                try:
                    # Try to Save the Episode
                    tmpEpisode.db_save()

                    # Now we need to move the entry in the Data Tree
                    # Start by building the Episode's new Node List
                    nodeData = (_('Libraries'), tmpLibrary.id, tmpEpisode.id)
                    # Add the Episode node to the data tree
                    treeCtrl.add_Node('EpisodeNode', nodeData, tmpEpisode.number, tmpLibrary.number, expandNode=False)
                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("AE %s >|< %s" % (nodeData[-2], nodeData[-1]))

                    # Now request that the Episode's OLD node be deleted.
                    # First build the Node List for the OLD Episode ...
                    nodeData = (_('Libraries'), oldLibrary.id, tmpEpisode.id)
                    # ... and delete it from the tree
                    treeCtrl.delete_Node(nodeData, 'EpisodeNode')

                    # If we are moving an Episode, the episode's Notes need to travel with the Episode.  The first step is to
                    # get a list of those Notes.
                    noteList = DBInterface.list_of_notes(Episode=tmpEpisode.number)
                    # If there are Episode Notes, we need to make sure they travel with the Episode
                    if noteList != []:
                        # Build the Node List for the new Episode
                        nodeData = (_('Libraries'), tmpLibrary.id, tmpEpisode.id)
                        # Select the new Episode Node
                        newNode = treeCtrl.select_Node(nodeData, 'EpisodeNode')
                        # Use the TreeCtrl's "add_note_nodes" method to move the notes locally
                        treeCtrl.add_note_nodes(noteList, newNode, Episode=tmpEpisode.number)
                        treeCtrl.Refresh()
                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            # Iterate through the Notes List
                            for noteid in noteList:
                                # Construct the message and data to be passed
                                msg = "AEN %s"
                                # To avoid problems in mixed-language environments, we need the UNTRANSLATED string here!
                                data = (u'Libraries',) + nodeData[1:]  + (noteid,)
                                # Build the message to be sent
                                for nd in data[1:]:
                                    msg += " >|< %s"
                                # Send the message
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage(msg % data)

                    # If we are moving an Episode, the episode's Transcripts need to travel with the Episode.  The first step is to
                    # get a list of those Transcripts.
                    transcriptList = DBInterface.list_transcripts(tmpLibrary.id, tmpEpisode.id)
                    # If there are Episode Transcripts, we need to make sure they travel with the Episode
                    if transcriptList != []:
                        # Iterate through the Transcript List
                        for (transcriptNum, transcriptID, transcriptEpisodeNum) in transcriptList:
                            # Build the Node List for the new Transcript
                            nodeData = (_('Libraries'), tmpLibrary.id, tmpEpisode.id, transcriptID)
                            # Add the Transcript node to the data tree
                            treeCtrl.add_Node('TranscriptNode', nodeData, transcriptNum, transcriptEpisodeNum, expandNode=False)
                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage("AT %s >|< %s >|< %s" % (nodeData[-3], nodeData[-2], nodeData[-1]))
                                    
                            # If we are moving a Transcript, the Transcript's Notes need to travel with the Transcript.  The first step is to
                            # get a list of those Notes.
                            noteList = DBInterface.list_of_notes(Transcript=transcriptNum)
                            
                            # If there are Episode Notes, we need to make sure they travel with the Episode
                            if noteList != []:
                                # Select the new Transcript Node
                                newNode = treeCtrl.select_Node(nodeData, 'TranscriptNode')
                                # Use the TreeCtrl's "add_note_nodes" method to move the notes locally
                                treeCtrl.add_note_nodes(noteList, newNode, Transcript=transcriptNum)
                                treeCtrl.Refresh()
                                # Now let's communicate with other Transana instances if we're in Multi-user mode
                                if not TransanaConstants.singleUserVersion:
                                    # Iterate through the Notes List
                                    for noteid in noteList:
                                        # Construct the message and data to be passed
                                        msg = "ATN %s"
                                        # To avoid problems in mixed-language environments, we need the UNTRANSLATED string here!
                                        data = (u'Libraries',) + nodeData[1:]  + (noteid,)
                                        # Build the message to be sent
                                        for nd in data[1:]:
                                            msg += " >|< %s"
                                        # Send the message
                                        if TransanaGlobal.chatWindow != None:
                                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                # If the Save fails ...
                except TransanaExceptions.SaveError, e:
                     # Display the Error Message
                     msg = _('An Episode named "%s" already exists in Library "%s".')
                     if 'unicode' in wx.PlatformInfo:
                         msg = unicode(msg, 'utf8')
                     errordlg = Dialogs.ErrorDialog(None, msg % (tmpEpisode.id, Library.Library(destNodeData.recNum).id))
                     errordlg.ShowModal()
                     errordlg.Destroy()
                    
                # If we get this far, unlock the Episode
                tmpEpisode.unlock_record()
            # If we are unable to lock the Episode ...
            except TransanaExceptions.RecordLockedError, e:
                # Report the Record Lock failure
                TransanaExceptions.ReportRecordLockedException(_('Episode'), tmpEpisode.id, e)
       
    # Drop a Collection on the ROOT Collection (Copy or Move a Collection to the root)
    # Drop a Collection on a Collection (Copy or Move a Collection to nest it in another collection)
    elif ((sourceData.nodetype == 'CollectionNode' and destNodeData.nodetype == 'CollectionsRootNode')) or \
         ((sourceData.nodetype == 'CollectionNode' and destNodeData.nodetype == 'CollectionNode')):
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceData.recNum, sourceData.parent)
        # If we're dropping on the Collections Root Node ...
        if destNodeData.nodetype == 'CollectionsRootNode':
            # ... create an EMPTY collection as the destination.  It has a number of 0, which is what it needs to have!
            destCollection = Collection.Collection()
            # Create a prompt for the user
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Collection "%s" \nto the main Collections node?'), 'utf8') % (copyMovePrompt, sourceCollection.id)
            else:
                prompt = _('Do you want to %s Collection "%s" \nto the main Collections node?') % (copyMovePrompt, sourceCollection.id)
        else:
            # Load the Destination Collection
            destCollection = Collection.Collection(destNodeData.recNum, destNodeData.parent)
            # We have to make sure the DESTINATION collection is not a child of the SOURCE collection here.
            # We can do this by comparing the collection paths, as spelled out in their Node Data
            if sourceCollection.GetNodeData() == destCollection.GetNodeData()[:len(sourceCollection.GetNodeData())]:
                # Create a prompt for the user
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("You cannot nest a Collection in one of its nested collections."), 'utf8')
                else:
                    prompt = _("You cannot nest a Collection in one of its nested collections.")
                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, prompt)
                errordlg.ShowModal()
                errordlg.Destroy()
                # We need to exit this method.
                return
            
            # Create a prompt for the user
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Collection "%s" \nso it is nested in Collection "%s"?'), 'utf8') % (copyMovePrompt, sourceCollection.id, destCollection.id)
            else:
                prompt = _('Do you want to %s Collection "%s" \nso it is nested in Collection "%s"?') % (copyMovePrompt, sourceCollection.id, destCollection.id)
        # Get user confirmation of the Collection Copy request
        dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
        result = dlg.LocalShowModal()
        dlg.Destroy()
        # If the user confirms ...
        if result == wx.ID_YES:
            # If a COPY is requested ...
            if action == 'Copy':
                # Start a List that will track collection being copied, so nested collections can be copied too
                collectionsToCopy = [(sourceCollection, destCollection)]
                # As long as there are more collections to copy ...
                while len(collectionsToCopy) > 0:
                    # ... note the source and destination collection information for THIS copy ...
                    (sourceCollection, destCollection) = collectionsToCopy[0]
                    # ... and remove the first entry 
                    collectionsToCopy = collectionsToCopy[1:]

                    # We need a list of all the quotes in the Source Collection
                    quoteList = DBInterface.list_of_quotes_by_collectionnum(sourceCollection.number)
                    # We need a list of all the clips in the Source Collection
                    clipList = DBInterface.list_of_clips_by_collection(sourceCollection.id, sourceCollection.parent)
                    # Start exception handling
                    try:
                        # Create a new collection
                        newCollection = Collection.Collection()
                        # Copy the source collection's information
                        newCollection.id = sourceCollection.id
                        # ... except that the parent collection should be the destination collection
                        newCollection.parent = destCollection.number
                        newCollection.comment = sourceCollection.comment
                        newCollection.owner = sourceCollection.owner
                        newCollection.keyword_group = sourceCollection.keyword_group
                        # Save the new collection.  This will throw a SaveError if it is a duplicate.
                        newCollection.db_save()
                        # Build the Node List for the tree control
                        nodeList = (_('Collections'), ) + newCollection.GetNodeData()
                        # Add the new Collection Node to the Tree
                        treeCtrl.add_Node('CollectionNode', nodeList, newCollection.number, newCollection.parent)
                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            # Prepare an Add Collection message
                            msg = "AC %s"
                            # Convert the Node List to the form needed for messaging
                            data = (nodeList[1],)
                            for nd in nodeList[2:]:
                                msg += " >|< %s"
                                data += (nd, )
                            # Send the message
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage(msg % data)

                        # Before we do anything else, let's add any nested collections to the list of collections to be processed.
                        # First, get a list of the current source collection's nested collections
                        nestedCollections = DBInterface.list_of_collections(sourceCollection.number)
                        # Iterate through the list of nested collections
                        for nestedColl in nestedCollections:
                            # Get the Collection Record for the nested collection
                            tempColl = Collection.Collection(nestedColl[0], nestedColl[2])
                            # Add the nested collection to the Copy list, nesting it in the current New collection
                            collectionsToCopy.append((tempColl, newCollection))

                        # Now it's time to copy the CLIPS. 
                        # Get the Tree Node for the collection we just added to the tree
                        newDestNode = treeCtrl.select_Node(nodeList, 'CollectionNode')

                        # We need to copy clips and snapshots in their MIXED order, rather than copying all Clips
                        # and then copying all Snapshots, as that would leave them in the wrong order in the destination,
                        # with all Clips followed by all Snapshots.

                        # Create a dictionary of objects which uses the sort_order as the key!!
                        copyOrder = {}

                        # Iterate through the list of Quotes
                        for quote in quoteList:
                            # Load the next Quote from the list
                            tempQuote = Quote.Quote(num=quote[0])
                            # Add the Quote to the sortOrder dictionary
                            copyOrder[tempQuote.sort_order] = tempQuote

                        # Iterate through the list of Clips
                        for clip in clipList:
                            # Load the next Clip from the list
                            tempClip = Clip.Clip(id_or_num=clip[0])
                            # Add the Clip to the sortOrder dictionary
                            copyOrder[tempClip.sort_order] = tempClip

                        # If Snapshots are supported ...
                        if TransanaConstants.proVersion:
                            # Now it's time to copy the SNAPSHOTS. 
                            # We need a list of all the snapshots in the Source Collection
                            snapshotList = DBInterface.list_of_snapshots_by_collectionnum(sourceCollection.number)
                            # Iterate through the list of Snapshots
                            for snapshot in snapshotList:
                                # Load the next Snapshot from the list
                                tempSnapshot = Snapshot.Snapshot(num_or_id=snapshot[0])
                                # Add the Snapshot to the sortOrder dictionary
                                copyOrder[tempSnapshot.sort_order] = tempSnapshot

                        # Get the dictionary keys
                        keys = copyOrder.keys()
                        # Sort the dictionary keys
                        keys.sort()
                        # Now iterate through the sorted keys ...
                        for key in keys:
                            # ... and get the next object that should go into the new destination
                            obj = copyOrder[key]
                            # If we have a Quote ...
                            if isinstance(obj, Quote.Quote):
                                # ... copy or Move the Quote to the Destination Collection
                                CopyMoveQuote(treeCtrl, newDestNode, obj, sourceCollection, newCollection, action)
                            # If we have a Clip ...
                            elif isinstance(obj, Clip.Clip):
                                # ... copy or Move the Clip to the Destination Collection
                                CopyMoveClip(treeCtrl, newDestNode, obj, sourceCollection, newCollection, action)
                            # If we have a Snapshot ...
                            elif isinstance(obj, Snapshot.Snapshot):
                                # ... copy or Move the Snapshot to the Destination Collection
                                CopyMoveSnapshot(treeCtrl, newDestNode, obj, sourceCollection, newCollection, action)

                        # See if the Keyword visualization needs to be updated.
                        treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
                        # Even if this computer doesn't need to update the keyword visualization others, might need to.
                        if not TransanaConstants.singleUserVersion:
                            # We need to update the Episode Keyword Visualization
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))

                        # Now let's copy the collection notes
                        notesList = DBInterface.list_of_notes(Collection=sourceCollection.number, includeNumber=True)
                        for note in notesList:
                            srcNote = Note.Note(note[0])
                            destNote = srcNote.duplicate()
                            destNote.collection_num = newCollection.number
                            destNote.db_save()
                            treeCtrl.add_Node('CollectionNoteNode', nodeList + (destNote.id,), destNote.number, newCollection.number)
                            # Now let's communicate with other Transana instances if we're in Multi-user mode
                            if not TransanaConstants.singleUserVersion:
                                # Construct the message and data to be passed
                                msg = "ACN %s"
                                # To avoid problems in mixed-language environments, we need the UNTRANSLATED string here!
                                data = ('Collections',)
                                for nd in (nodeList[1:] + (destNote.id,)):
                                    msg += " >|< %s"
                                    data += (nd, )
                                if TransanaGlobal.chatWindow != None:
                                    TransanaGlobal.chatWindow.SendMessage(msg % data)

                    # Handle "SaveError" exception
                    except TransanaExceptions.SaveError:
                        # Get the error message
                        msg = sys.exc_info()[1].reason
                        # Get the comparison prompt for Collections
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            comparePrompt = unicode(_('A Collection named "%s" already exists.\nPlease enter a different Collection ID.'), 'utf8')
                        else:
                            comparePrompt = _('A Collection named "%s" already exists.\nPlease enter a different Collection ID.')
                        # If we have a Duplicate Collection error ...
                        if ('\n' in msg) and (msg.split('\n')[1] == comparePrompt.split('\n')[1]):
                            # ... only print the first part of the message!
                            msg = msg.split('\n')[0]
                        # Display the Error Message
                        errordlg = Dialogs.ErrorDialog(None, msg)
                        errordlg.ShowModal()
                        errordlg.Destroy()
                    # Handle all other exceptions
                    except:
                        if DEBUG:
                            import traceback
                            traceback.print_exc(file=sys.stdout)

                        # Display the Exception Message
                        prompt = "%s : %s"
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(prompt, 'utf8')
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                                    
            # If a MOVE is requested ...
            elif action == 'Move':
                # Start Exception handling
                try:
                    # Remember the node list for the original source node  
                    originalNodeList = (_('Collections'),) + sourceCollection.GetNodeData()
                    # Lock the source collection, if possible.
                    sourceCollection.lock_record()
                    # Change the Collection Parent to the NEW parent collection
                    sourceCollection.parent = destCollection.number
                    # Save the changes
                    sourceCollection.db_save()
                    # Unlock the Collection
                    sourceCollection.unlock_record()
                    # Now add the new Tree Node.  Build the Node List for the tree control
                    nodeList = (_('Collections'), ) + sourceCollection.GetNodeData()
                    # Move the node, signalling that MU messages SHOULD be sent.
                    treeCtrl.copy_Node('CollectionNode', originalNodeList, nodeList[:-1], True, sendMessage=True)

                # Handle "SaveError" exception
                except TransanaExceptions.SaveError:
                    # Get the error message
                    msg = sys.exc_info()[1].reason
                    # Get the comparison prompt for Collections
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        comparePrompt = unicode(_('A Collection named "%s" already exists.\nPlease enter a different Collection ID.'), 'utf8')
                    else:
                        comparePrompt = _('A Collection named "%s" already exists.\nPlease enter a different Collection ID.')
                    # If we have a Duplicate Collection error ...
                    if ('\n' in msg) and (msg.split('\n')[1] == comparePrompt.split('\n')[1]):
                        # ... only print the first part of the message!
                        msg = msg.split('\n')[0]
                    # Display the Error Message
                    errordlg = Dialogs.ErrorDialog(None, msg)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                # Handle all other exceptions
                except:
                  if DEBUG:
                      import traceback
                      traceback.print_exc(file=sys.stdout)

                  # Display the Exception Message
                  prompt = "%s : %s"
                  if 'unicode' in wx.PlatformInfo:
                      # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                      prompt = unicode(prompt, 'utf8')
                  errordlg = Dialogs.ErrorDialog(None, prompt % (sys.exc_info()[0],sys.exc_info()[1]))
                  errordlg.ShowModal()
                  errordlg.Destroy()
              
        # See if the Keyword visualization needs to be updated.
        treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
        # Even if this computer doesn't need to update the keyword visualization others, might need to.
        if not TransanaConstants.singleUserVersion:
            # We need to update the Episode Keyword Visualization
            if TransanaGlobal.chatWindow != None:
                TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))

    # Drop a Quote on a Collection (Copy or Move a Quote)
    elif (sourceData.nodetype == 'QuoteNode' and destNodeData.nodetype == 'CollectionNode'):
        # Load the Source Quote
        sourceQuote = Quote.Quote(num=sourceData.recNum)
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceData.parent)
        # Load the Destination Collection
        destCollection = Collection.Collection(destNodeData.recNum, destNodeData.parent)

        # If "Yes to All" has not already been selected ...
        if not YESTOALL:
            # Get user confirmation of the Clip Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Quote "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceQuote.id, sourceCollection.id, destCollection.id)
            else:
                prompt = _('Do you want to %s Quote "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceQuote.id, sourceCollection.id, destCollection.id)
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt, yesToAll=True)
            result = dlg.LocalShowModal()
            # If the user selected Yes To All, we need to process that before destroying the Dialog
            if result == dlg.YESTOALLID:
                # Set the global YesToAll to True
                YESTOALL = True
                # Yes to All is a Yes.
                result = wx.ID_YES
            dlg.Destroy()
        else:
            result = wx.ID_YES
        
        if result == wx.ID_YES:
            try:
                # Copy or Move the Quote to the Destination Collection
                CopyMoveQuote(treeCtrl, destNode, sourceQuote, sourceCollection, destCollection, action)
                # See if the Keyword visualization needs to be updated.
                treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Episode Keyword Visualization
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Quote', sourceQuote.number, sourceQuote.source_document_num))
            except TransanaExceptions.RecordLockedError, e:
                TransanaExceptions.ReportRecordLockedException(_("Quote"), sourceQuote.id, e)
                
    # Drop a Quote on a Quote, Clip, or Snapshot (Alter SortOrder, Copy or Move a Quote to a particular place in the SortOrder)
    elif (sourceData.nodetype == 'QuoteNode' and destNodeData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']):
        # Load the Source Quote
        sourceQuote = Quote.Quote(num=sourceData.recNum)
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceData.parent)
        if destNodeData.nodetype == 'QuoteNode':
            # Load the Destination (or target) Quote
            destObj = Quote.Quote(num=destNodeData.recNum)
        elif destNodeData.nodetype == 'ClipNode':
            # Load the Destination (or target) Clip
            destObj = Clip.Clip(destNodeData.recNum)
        elif destNodeData.nodetype == 'SnapshotNode':
            # Load the Destination (or target) Snapshot
            destObj = Snapshot.Snapshot(destNodeData.recNum)
        # Load the Destination (or target) Collection
        destCollection = Collection.Collection(destObj.collection_num)
        # See if we are in the SAME Collection, and therefore just changing Sort Order
        if sourceCollection.number == destCollection.number:
            # If so, change the Sort Order as requested
            if not ChangeClipOrder(treeCtrl, destNode, sourceQuote, sourceCollection):
                # This is OKAY.  Nothing needs to be fixed, either locally or remotely if this fails.
                pass
            treeCtrl.UpdateCollectionSortOrder(treeCtrl.GetItemParent(destNode))
            wx.CallAfter(treeCtrl.Refresh)

        # If not, we are copying/moving a Quote to a place in the SortOrder in a Different Colletion
        else:

            # If "Yes to All" has not already been selected ...
            if not YESTOALL:
                # Get user confirmation of the Quote Copy/Move request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Do you want to %s Quote "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceQuote.id, sourceCollection.id, destCollection.id)
                else:
                    prompt = _('Do you want to %s Quote "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceQuote.id, sourceCollection.id, destCollection.id)
                dlg = Dialogs.QuestionDialog(treeCtrl, prompt, yesToAll=True)
                result = dlg.LocalShowModal()
                # If the user selected Yes To All, we need to process that before destroying the Dialog
                if result == dlg.YESTOALLID:
                    # Set YesToAll to True
                    YESTOALL = True
                    # YesToAll is a Yes
                    result = wx.ID_YES
                dlg.Destroy()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                try:
                    # Copy or Move the Quote to the Destination Collection
                    # If confirmed, copy the Source Quote to the Destination Collection.  CopyMoveQuote will place the Quote at
                    # end of the list of quotes, clips, and snapshots.
                    # We need to work with the COPY of the Quote instead of the original from here on, so we get that
                    # value from CopyQuote.
                    tempQuote = CopyMoveQuote(treeCtrl, destNode, sourceQuote, sourceCollection, destCollection, action)
                    # If the Copy/Move is cancelled, tempQuote will be None
                    if tempQuote != None:
                        # Now change the order of the Quotes
                        if not ChangeClipOrder(treeCtrl, destNode, tempQuote, destCollection):
                            # This is OKAY.  Nothing needs to be fixed, either locally or remotely if this fails.
                            pass
                            
                    treeCtrl.UpdateCollectionSortOrder(treeCtrl.GetItemParent(destNode))
                    wx.CallAfter(treeCtrl.Refresh)

                    # See if the Keyword visualization needs to be updated.
                    treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
                    # Even if this computer doesn't need to update the keyword visualization others, might need to.
                    if not TransanaConstants.singleUserVersion:
                        # We need to update the Document Keyword Visualization
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))
                except TransanaExceptions.RecordLockedError, e:
                    TransanaExceptions.ReportRecordLockedException(_("Quote"), sourceQuote.id, e)

    # Drop a Clip on a Collection (Copy or Move a Clip)
    elif (sourceData.nodetype == 'ClipNode' and destNodeData.nodetype == 'CollectionNode'):
        # Load the Source Clip
        sourceClip = Clip.Clip(id_or_num=sourceData.recNum)
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceData.parent)
        # Load the Destination Collection
        destCollection = Collection.Collection(destNodeData.recNum, destNodeData.parent)

        # If "Yes to All" has not already been selected ...
        if not YESTOALL:
            # Get user confirmation of the Clip Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
            else:
                prompt = _('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt, yesToAll=True)
            result = dlg.LocalShowModal()
            # If the user selected Yes To All, we need to process that before destroying the Dialog
            if result == dlg.YESTOALLID:
                # Set the global YesToAll to True
                YESTOALL = True
                # Yes to All is a Yes.
                result = wx.ID_YES
            dlg.Destroy()
        else:
            result = wx.ID_YES
        
        if result == wx.ID_YES:
            try:
                # Copy or Move the Clip to the Destination Collection
                CopyMoveClip(treeCtrl, destNode, sourceClip, sourceCollection, destCollection, action)
                # See if the Keyword visualization needs to be updated.
                treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Episode Keyword Visualization
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', sourceClip.number, sourceClip.episode_num))
            except TransanaExceptions.RecordLockedError, e:
                TransanaExceptions.ReportRecordLockedException(_("Clip"), sourceClip.id, e)
                
    # Drop a Clip on a Clip (Alter SortOrder, Copy or Move a Clip to a particular place in the SortOrder)
    elif (sourceData.nodetype == 'ClipNode' and destNodeData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']):
        # Load the Source Clip
        sourceClip = Clip.Clip(id_or_num=sourceData.recNum)
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceData.parent)
        if destNodeData.nodetype == 'QuoteNode':
            # Load the Destination (or target) Quote
            destObj = Quote.Quote(num=destNodeData.recNum)
        elif destNodeData.nodetype == 'ClipNode':
            # Load the Destination (or target) Clip
            destObj = Clip.Clip(destNodeData.recNum)
        elif destNodeData.nodetype == 'SnapshotNode':
            # Load the Destination (or target) Snapshot
            destObj = Snapshot.Snapshot(destNodeData.recNum)
        # Load the Destination (or target) Collection
        destCollection = Collection.Collection(destObj.collection_num)
        # See if we are in the SAME Collection, and therefore just changing Sort Order
        if sourceCollection.number == destCollection.number:
            # If so, change the Sort Order as requested
            if not ChangeClipOrder(treeCtrl, destNode, sourceClip, sourceCollection):
                # This is OKAY.  Nothing needs to be fixed, either locally or remotely if this fails.
                pass
            treeCtrl.UpdateCollectionSortOrder(treeCtrl.GetItemParent(destNode))
            wx.CallAfter(treeCtrl.Refresh)

        # If not, we are copying/moving a Clip to a place in the SortOrder in a Different Colletion
        else:

            # If "Yes to All" has not already been selected ...
            if not YESTOALL:
                # Get user confirmation of the Clip Copy/Move request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
                else:
                    prompt = _('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
                dlg = Dialogs.QuestionDialog(treeCtrl, prompt, yesToAll=True)
                result = dlg.LocalShowModal()
                # If the user selected Yes To All, we need to process that before destroying the Dialog
                if result == dlg.YESTOALLID:
                    # Set YesToAll to True
                    YESTOALL = True
                    # YesToAll is a Yes
                    result = wx.ID_YES
                dlg.Destroy()
            else:
                result = wx.ID_YES

            if result == wx.ID_YES:
                try:
                    # Copy or Move the Clip to the Destination Collection
                    # If confirmed, copy the Source Clip to the Destination Collection.  CopyClip will place the clip at
                    # end of the list of clips.
                    # We need to work with the COPY of the clip instead of the original from here on, so we get that
                    # value from CopyClip.
                    tempClip = CopyMoveClip(treeCtrl, destNode, sourceClip, sourceCollection, destCollection, action)
                    # If the Copy/Move is cancelled, tempClip will be None
                    if tempClip != None:
                        # Now change the order of the clips
                        if not ChangeClipOrder(treeCtrl, destNode, tempClip, destCollection):
                            # This is OKAY.  Nothing needs to be fixed, either locally or remotely if this fails.
                            pass
                            
                    treeCtrl.UpdateCollectionSortOrder(treeCtrl.GetItemParent(destNode))
                    wx.CallAfter(treeCtrl.Refresh)

                    # See if the Keyword visualization needs to be updated.
                    treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
                    # Even if this computer doesn't need to update the keyword visualization others, might need to.
                    if not TransanaConstants.singleUserVersion:
                        # We need to update the Episode Keyword Visualization
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))
                except TransanaExceptions.RecordLockedError, e:
                    TransanaExceptions.ReportRecordLockedException(_("Clip"), sourceClip.id, e)

    # Drop a Clip on a Keyword (Create Keyword Example)
    elif (sourceData.nodetype == 'ClipNode' and destNodeData.nodetype == 'KeywordNode'):
        # Get the Keyword Group 
        kwg = treeCtrl.GetItemText(treeCtrl.GetItemParent(destNode))
        # Get the Keyword
        kw = treeCtrl.GetItemText(destNode)
        # Get user confirmation of the Keyword Example Add request
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to add Clip "%s" as an example of Keyword "%s:%s"?'), 'utf8') % (sourceData.text, kwg, kw)
        else:
            prompt = _('Do you want to add Clip "%s" as an example of Keyword "%s:%s"?') % (sourceData.text, kwg, kw)
        dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
        result = dlg.LocalShowModal()
        dlg.Destroy()
        if result == wx.ID_YES:
            # Locate the appropriate ClipKeyword Record in the Database and flag it as an example
            DBInterface.SetKeywordExampleStatus(kwg, kw, sourceData.recNum, 1)
            # Add the Clip as Keyword Example to the Node Structure.
            # (Keyword Root, Keyword Group, Keyword, Example Clip Name)
            nodeData = (_('Keywords'), kwg, kw, sourceData.text)
            # Add the Keyword Example Node to the Database Tree Tab
            treeCtrl.add_Node("KeywordExampleNode", nodeData, sourceData.recNum, sourceData.parent)

            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion:
                # The first message parameter for a Keyword Example is the Clip Number
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage("AKE %d >|< %s >|< %s >|< %s" % (sourceData.recNum, kwg, kw, sourceData.text))
                    TransanaGlobal.chatWindow.SendMessage("UKL Clip %d" % sourceData.recNum)

            # Load the Clip
            tempClip = Clip.Clip(sourceData.recNum)
            # If the affected clip is for the current Episode, we need to update the
            # Keyword Visualization
            if (isinstance(treeCtrl.parent.ControlObject.currentObj, Episode.Episode)) and \
                (tempClip.episode_num == treeCtrl.parent.ControlObject.currentObj.number):
                # See if the Keyword visualization needs to be updated.
                treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
            # If the affected clip is the current Clip, we need to update the
            # Keyword Visualization
            if (isinstance(treeCtrl.parent.ControlObject.currentObj, Clip.Clip)) and \
               (sourceData.recNum == treeCtrl.parent.ControlObject.currentObj.number):
                # See if the Keyword visualization needs to be updated.
                treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
            # Even if this computer doesn't need to update the keyword visualization others, might need to.
            if not TransanaConstants.singleUserVersion:
                # We need to update the Clip Keyword Visualization when adding a keyword to a clip
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', tempClip.number, tempClip.episode_num))

    # Drop a Snapshot on a Collection (Copy or Move a Snapshot)
    elif (sourceData.nodetype == 'SnapshotNode' and destNodeData.nodetype == 'CollectionNode'):
        # Load the Source Snapshot
        sourceSnapshot = Snapshot.Snapshot(sourceData.recNum)
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceData.parent)
        # Load the Destination Collection
        destCollection = Collection.Collection(destNodeData.recNum, destNodeData.parent)

        # If "Yes to All" has not already been selected ...
        if not YESTOALL:
            # Get user confirmation of the Clip Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Snapshot "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceSnapshot.id, sourceCollection.id, destCollection.id)
            else:
                prompt = _('Do you want to %s Snapshot "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceSnapshot.id, sourceCollection.id, destCollection.id)
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt, yesToAll=True)
            result = dlg.LocalShowModal()
            # If the user selected Yes To All, we need to process that before destroying the Dialog
            if result == dlg.YESTOALLID:
                # Set YesToAll to True
                YESTOALL = True
                # YesToAll is a Yes
                result = wx.ID_YES
            dlg.Destroy()
        # If we have YesToAll ...
        else:
            # ... then that's a Yes
            result = wx.ID_YES

        if result == wx.ID_YES:
            try:
                # Copy or Move the Snapshot to the Destination Collection
                CopyMoveSnapshot(treeCtrl, destNode, sourceSnapshot, sourceCollection, destCollection, action)
                # See if the Keyword visualization needs to be updated.
                treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if (sourceSnapshot.episode_num > 0) and not TransanaConstants.singleUserVersion:
                    # We need to update the Episode Keyword Visualization
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Episode', sourceSnapshot.episode_num, 0)) # ('Clip', sourceClip.number, sourceClip.episode_num))
            except TransanaExceptions.RecordLockedError, e:
                TransanaExceptions.ReportRecordLockedException(_("Snapshot"), sourceSnapshot.id, e)
                
    # Drop a Snapshot on a Quote, Clip or Snapshot (Alter SortOrder, Copy or Move a Snapshot to a particular place in the SortOrder)
    elif (sourceData.nodetype == 'SnapshotNode' and destNodeData.nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']):
        # Load the Source Snapshot
        sourceSnapshot = Snapshot.Snapshot(sourceData.recNum)
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceData.parent)
        if destNodeData.nodetype == 'QuoteNode':
            # Load the Destination (or target) Quote
            destObj = Quote.Quote(num=destNodeData.recNum)
        elif destNodeData.nodetype == 'ClipNode':
            # Load the Destination (or target) Clip
            destObj = Clip.Clip(destNodeData.recNum)
        elif destNodeData.nodetype == 'SnapshotNode':
            # Load the Destination (or target) Snapshot
            destObj = Snapshot.Snapshot(destNodeData.recNum)
        # Load the Destination (or target) Collection
        destCollection = Collection.Collection(destObj.collection_num)
        # See if we are in the SAME Collection, and therefore just changing Sort Order
        if sourceCollection.number == destCollection.number:
            # If so, change the Sort Order as requested
            if not ChangeClipOrder(treeCtrl, destNode, sourceSnapshot, sourceCollection):
                # This is OKAY.  Nothing needs to be fixed, either locally or remotely if this fails.
                pass
            treeCtrl.UpdateCollectionSortOrder(treeCtrl.GetItemParent(destNode))
            wx.CallAfter(treeCtrl.Refresh)

        # If not, we are copying/moving a Snapshot to a place in the SortOrder in a different Collection ...
        else:

            # If "Yes to All" has not already been selected ...
            if not YESTOALL:
                # Get user confirmation of the Snapshot Copy/Move request
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Do you want to %s Snapshot "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceSnapshot.id, sourceCollection.id, destCollection.id)
                else:
                    prompt = _('Do you want to %s Snapshot "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceSnapshot.id, sourceCollection.id, destCollection.id)
                dlg = Dialogs.QuestionDialog(treeCtrl, prompt, yesToAll=True)
                result = dlg.LocalShowModal()
                # If the user selected Yes To All, we need to process that before destroying the Dialog
                if result == dlg.YESTOALLID:
                    # Set YesToAll to True
                    YESTOALL = True
                    # YesToAll is a Yes
                    result = wx.ID_YES
                dlg.Destroy()
            # If we have YesToAll ...
            else:
                # ... then that's a Yes
                result = wx.ID_YES

            if result == wx.ID_YES:
                try:
                    # Copy or Move the Snapshot to the Destination Collection
                    # If confirmed, copy the Source Snapshot to the Destination Collection.  CopyMoveSnapshot will place the Snapshot at
                    # end of the list of objects.
                    # We need to work with the COPY of the snapshot instead of the original from here on, so we get that
                    # value from CopyMoveSnapshot.
                    tempObject = CopyMoveSnapshot(treeCtrl, destNode, sourceSnapshot, sourceCollection, destCollection, action)
                    # If the Copy/Move is cancelled, tempObject will be None
                    if tempObject != None:
                        # Now change the order of the objects
                        if not ChangeClipOrder(treeCtrl, destNode, tempObject, destCollection):
                            # This is OKAY.  Nothing needs to be fixed, either locally or remotely if this fails.
                            pass

                    treeCtrl.UpdateCollectionSortOrder(treeCtrl.GetItemParent(destNode))
                    wx.CallAfter(treeCtrl.Refresh)
                            
                except TransanaExceptions.RecordLockedError, e:
                    TransanaExceptions.ReportRecordLockedException(_("Snapshot"), sourceSnapshot.id, e)

    # Drop a Keyword on a Library
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'LibraryNode'):
        DropKeyword(treeCtrl, sourceData, 'Libraries', treeCtrl.GetItemText(destNode), destNodeData.recNum, 0, confirmations=confirmations)
   
    # Drop a Keyword on a Document
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'DocumentNode'):
        DropKeyword(treeCtrl, sourceData, 'Document', treeCtrl.GetItemText(destNode), destNodeData.recNum, 0, confirmations=confirmations)
   
    # Drop a Keyword on an Episode
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'EpisodeNode'):
        DropKeyword(treeCtrl, sourceData, 'Episode', treeCtrl.GetItemText(destNode), destNodeData.recNum, 0, confirmations=confirmations)
   
    # Drop a Keyword on a Collection
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'CollectionNode'):
        DropKeyword(treeCtrl, sourceData, 'Collection', treeCtrl.GetItemText(destNode), destNodeData.recNum, destNodeData.parent, confirmations=confirmations)

    # Drop a Keyword on a Quote
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'QuoteNode'):
        DropKeyword(treeCtrl, sourceData, 'Quote', treeCtrl.GetItemText(destNode), destNodeData.recNum, 0, confirmations=confirmations)

    # Drop a Keyword on a Clip
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'ClipNode'):
        DropKeyword(treeCtrl, sourceData, 'Clip', treeCtrl.GetItemText(destNode), destNodeData.recNum, 0, confirmations=confirmations)

    # Drop a Keyword on a Snapshot
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'SnapshotNode'):
        DropKeyword(treeCtrl, sourceData, 'Snapshot', treeCtrl.GetItemText(destNode), destNodeData.recNum, 0, confirmations=confirmations)

    # Drop a Keyword on a Keyword Group (Copy or Move a Keyword)
    elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'KeywordGroupNode'):
        # Get user confirmation of the Keyword Copy request
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to %s Keyword "%s" from\nKeyword Group "%s" to\nKeyword Group "%s"?'), 'utf8') % (copyMovePrompt, sourceData.text, sourceData.parent, treeCtrl.GetItemText(destNode))
        else:
            prompt = _('Do you want to %s Keyword "%s" from\nKeyword Group "%s" to\nKeyword Group "%s"?') % (copyMovePrompt, sourceData.text, sourceData.parent, treeCtrl.GetItemText(destNode))
        dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
        result = dlg.LocalShowModal()
        dlg.Destroy()
        if result == wx.ID_YES:
            # Determine if we're copying or moving data
            if action == 'Copy':
                # If confirmed, copy the Keyword to the Keyword Group
                # Load the original Keyword
                originalKeyword = Keyword.Keyword(sourceData.parent, sourceData.text)
                # Create a blank Keyword Object
                tempKeyword = Keyword.Keyword()
                # Assign the data value for the Keyword
                tempKeyword.keyword = sourceData.text
                # Copy the keyword definition
                tempKeyword.definition = originalKeyword.definition
                # Copy the Line Color
                tempKeyword.lineColorName = originalKeyword.lineColorName
                tempKeyword.lineColorDef = originalKeyword.lineColorDef
                # Copy the Draw Mode
                tempKeyword.drawMode = originalKeyword.drawMode
                # Copy the Line Width
                tempKeyword.lineWidth = originalKeyword.lineWidth
                # Copy the Line Style
                tempKeyword.lineStyle = originalKeyword.lineStyle
            elif action == 'Move':
                # If confirmed, move the Keyword to the Clip.
                # Load the appropriate Keyword Object
                tempKeyword = Keyword.Keyword(sourceData.parent, sourceData.text)
                # Lock the Keyword Record
                tempKeyword.lock_record()
                # Define the Source Node List
                sourceData.nodeList = (_('Keywords'), sourceData.parent, sourceData.text)
                
                # We need to find out if the moved keyword had any keyword Example nodes.
                # First, we need to get the original Keyword node.
                sourceNode = treeCtrl.select_Node(sourceData.nodeList, sourceData.nodetype)
                # Create a List to put information about Examples in.
                exampleNodes = []
                # Check to see if the node has children ...
                if treeCtrl.ItemHasChildren(sourceNode):
                    # ... and if it does, get the first.
                    (childNode, cookieVal) = treeCtrl.GetFirstChild(sourceNode)
                    # As long as the node is OK ...
                    while childNode.IsOk():
                        # ... get the node's Data ...
                        childData = treeCtrl.GetPyData(childNode)
                        # ... append the node's name and data to the Example Nodes list ...
                        exampleNodes.append((treeCtrl.GetItemText(childNode), childData))
                        # ... and get the next child, if there is one.
                        (childNode, cookieVal) = treeCtrl.GetNextChild(sourceNode, cookieVal)

            # Change the Keyword Group to the new Group Value (from dropNode)
            tempKeyword.keywordGroup = treeCtrl.GetItemText(destNode)
            try:
                # Save the Keyword Record
                tempKeyword.db_save()
                if action == 'Move':
                    # Unlock the Keyword Record
                    tempKeyword.unlock_record()
                    # Remove the old Keyword from the Tree.  First, build the appropriate Node List
                    nodeList = (_('Keywords'), sourceData.parent, sourceData.text)
                    # Now delete the specified node
                    treeCtrl.delete_Node(nodeList, 'KeywordNode')
                    # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                    ClearClipboard()

                # Add the new Keyword to the Tree.  First, build the Node List as required by add_Node()
                nodeList = (_('Keywords'), treeCtrl.GetItemText(destNode), sourceData.text)
                # Now call add_Node to actually put the node in the tree.
                treeCtrl.add_Node('KeywordNode', nodeList, 0, treeCtrl.GetItemText(destNode))

                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("AK %s >|< %s" % (treeCtrl.GetItemText(destNode), sourceData.text))

                # If we are Moving and if there were Examples, we need to add those examples to the moved Keyword Node.
                if (action == 'Move') and (exampleNodes != []):
                    # Iterate through the Example nodes data collected above.
                    for (nodeName, nodeData) in exampleNodes:
                        # Add the Keyword Example to the Tree.  First, build the Node List as required by add_Node()
                        nodeList = (_('Keywords'), treeCtrl.GetItemText(destNode), sourceData.text, nodeName)
                        # Now call add_Node to actually put the node in the tree.
                        treeCtrl.add_Node("KeywordExampleNode", nodeList, nodeData.recNum, nodeData.parent)

                        # Now let's communicate with other Transana instances if we're in Multi-user mode
                        if not TransanaConstants.singleUserVersion:
                            # The first message parameter for a Keyword Example is the Clip Number
                            if TransanaGlobal.chatWindow != None:
                                TransanaGlobal.chatWindow.SendMessage("AKE %d >|< %s >|< %s >|< %s" % (nodeData.recNum, treeCtrl.GetItemText(destNode), sourceData.text, nodeName))

                # If it's a Move, we need to update the Keyword Visualization too!
                if (action == 'Move'):
                    # See if the Keyword visualization needs to be updated.
                    treeCtrl.parent.ControlObject.UpdateKeywordVisualization()
                    # Even if this computer doesn't need to update the keyword visualization others, might need to.
                    if not TransanaConstants.singleUserVersion:
                        # We need to update the Keyword Visualization
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('None', 0, 0))

            except TransanaExceptions.SaveError:
                # We have locked this record.  We better unlock it.
                tempKeyword.unlock_record()
                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()

    # Drop a Library Note on a new Library
    elif (sourceData.nodetype == 'LibraryNoteNode' and destNodeData.nodetype == 'LibraryNode'):
        # Load the Source Library Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Library
        sourceLibrary = Library.Library(sourceNote.series_num)
        # Load the Destination Library
        destLibrary = Library.Library(destNodeData.recNum)
        # Can't drop a note on it's own parent!
        if sourceLibrary.number != destLibrary.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nLibrary "%s" to\nLibrary "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceLibrary.id, destLibrary.id)
            else:
                prompt = _('Do you want to %s Note "%s" from\nLibrary "%s" to\nLibrary "%s"?') % (copyMovePrompt, sourceNote.id, sourceLibrary.id, destLibrary.id)
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Library
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Library, alter its Library Number
                    newNote.series_num = destLibrary.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Library Number
                        sourceNote.series_num = destLibrary.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Library "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destLibrary.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.  Start with the sourceLibrary.
                            nodeList = (_('Libraries'), sourceLibrary.id, sourceNote.id)
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'LibraryNoteNode')

                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Library "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destLibrary.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Clip Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Libraries'), destLibrary.id, tempNote.id)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('LibraryNoteNode', nodeList, tempNote.number, tempNote.series_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Library Note message
                        msg = "ASN Series >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop a Document Note on a new Document
    elif (sourceData.nodetype == 'DocumentNoteNode' and destNodeData.nodetype == 'DocumentNode'):
        # Load the Source Document Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Document
        sourceDocument = Document.Document(sourceNote.document_num)
        # Load the Destination Document
        destDocument = Document.Document(destNodeData.recNum)
        # Load the Destination Library
        destLibrary = Library.Library(destDocument.library_num)
        # Can't drop a note on it's own self!
        if sourceDocument.number != destDocument.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nDocument "%s" to\nDocument "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceDocument.id, destDocument.id)
            else:
                prompt = _('Do you want to %s Note "%s" from\nDocument "%s" to\nDocument "%s"?') % (copyMovePrompt, sourceNote.id, sourceDocument.id, destDocument.id)
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Document
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Document, alter its Document Number
                    newNote.document_num = destDocument.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Document Number
                        sourceNote.document_num = destDocument.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Document "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destDocument.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.  Start with the sourceLibrary.
                            nodeList = (_('Libraries'), destLibrary.id, sourceDocument.id, sourceNote.id)
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'DocumentNoteNode')
                                 
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Document "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destDocument.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Clip Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Libraries'), destLibrary.id, destDocument.id, tempNote.id)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('DocumentNoteNode', nodeList, tempNote.number, tempNote.document_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Document Note message
                        msg = "ADN Libraries >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop an Episode Note on a new Episode
    elif (sourceData.nodetype == 'EpisodeNoteNode' and destNodeData.nodetype == 'EpisodeNode'):
        # Load the Source Episode Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Episode
        sourceEpisode = Episode.Episode(sourceNote.episode_num)
        # Load the Destination Episode
        destEpisode = Episode.Episode(destNodeData.recNum)
        # Can't drop a note on it's own parent!
        if sourceEpisode.number != destEpisode.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nEpisode "%s" to\nEpisode "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceEpisode.id, destEpisode.id)
            else:
                prompt = _('Do you want to %s Note "%s" from\nEpisode "%s" to\nEpisode "%s"?') % (copyMovePrompt, sourceNote.id, sourceEpisode.id, destEpisode.id)
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Library
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Episode, alter its Episode Number
                    newNote.episode_num = destEpisode.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Episode Number
                        sourceNote.episode_num = destEpisode.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Episode "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destEpisode.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.  Start with the sourceLibrary.
                            nodeList = (_('Libraries'), sourceEpisode.series_id, sourceEpisode.id, sourceNote.id)
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'EpisodeNoteNode')
                                 
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Episode "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destEpisode.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Clip Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Libraries'), destEpisode.series_id, destEpisode.id, tempNote.id)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('EpisodeNoteNode', nodeList, tempNote.number, tempNote.episode_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Episode Note message
                        msg = "AEN Series >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop a Transcript Note on a new Transcript
    elif (sourceData.nodetype == 'TranscriptNoteNode' and destNodeData.nodetype == 'TranscriptNode'):
        # Load the Source Transcript Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Transcript
        # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
        sourceTranscript = Transcript.Transcript(sourceNote.transcript_num, skipText=True)
        # Load the Source Episode
        sourceEpisode = Episode.Episode(sourceTranscript.episode_num)
        # Load the Destination Transcript
        # To save time here, we can skip loading the actual transcript text, which can take time once we start dealing with images!
        destTranscript = Transcript.Transcript(destNodeData.recNum, skipText=True)
        # Load the Destination Episode
        destEpisode = Episode.Episode(destTranscript.episode_num)
        # Can't drop a note on it's own parent!
        if sourceTranscript.number != destTranscript.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nTranscript "%s" to\nTranscript "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceTranscript.id, destTranscript.id)
            else:
                prompt = _('Do you want to %s Note "%s" from\nTranscript "%s" to\nTranscript "%s"?') % (copyMovePrompt, sourceNote.id, sourceTranscript.id, destTranscript.id)
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Library
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Transcript, alter its Transcript Number
                    newNote.transcript_num = destTranscript.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Transcript Number
                        sourceNote.transcript_num = destTranscript.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Transcript "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destTranscript.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.  Start with the sourceLibrary.
                            nodeList = (_('Libraries'), sourceEpisode.series_id, sourceEpisode.id, sourceTranscript.id, sourceNote.id)
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'TranscriptNoteNode')

                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Transcript "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destTranscript.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Clip Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Libraries'), destEpisode.series_id, destEpisode.id, destTranscript.id, tempNote.id)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('TranscriptNoteNode', nodeList, tempNote.number, tempNote.transcript_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Transcript Note message
                        msg = "ATN Series >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                             msg += " >|< %s"
                             data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop a Collection Note on a new Collection
    elif (sourceData.nodetype == 'CollectionNoteNode' and destNodeData.nodetype == 'CollectionNode'):
        # Load the Source Collection Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Collection
        sourceCollection = Collection.Collection(sourceNote.collection_num)
        # Load the Destination Collection
        destCollection = Collection.Collection(destNodeData.recNum)
        # Can't drop a note on it's own parent!
        if sourceCollection.number != destCollection.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceCollection.GetNodeString(), destCollection.GetNodeString())
            else:
                prompt = _('Do you want to %s Note "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceNote.id, sourceCollection.GetNodeString(), destCollection.GetNodeString())
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Library
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Collection, alter its Collection Number
                    newNote.collection_num = destCollection.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Collection Number
                        sourceNote.collection_num = destCollection.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Collection "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destCollection.GetNodeString()))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.
                            nodeList = (_('Collections'),) + sourceCollection.GetNodeData() + (sourceNote.id, )
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'CollectionNoteNode')
                                 
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Collection "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destCollection.GetNodeString()))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Clip Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Collections'),) + destCollection.GetNodeData() + (tempNote.id,)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('CollectionNoteNode', nodeList, tempNote.number, tempNote.collection_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Collection Note message
                        msg = "ACN Collections >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                             msg += " >|< %s"
                             data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop a Quote Note on a new Quote
    elif (sourceData.nodetype == 'QuoteNoteNode' and destNodeData.nodetype == 'QuoteNode'):
        # Load the Source Quote Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Quote
        sourceQuote = Quote.Quote(num=sourceNote.quote_num)
        # Load the Destination Quote
        destQuote = Quote.Quote(destNodeData.recNum)
        # Can't drop a note on it's own parent!
        if sourceQuote.number != destQuote.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nQuote "%s" to\nQuote "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceQuote.id, destQuote.id)
            else:
                prompt = _('Do you want to %s Note "%s" from\nQuote "%s" to\nQuote "%s"?') % (copyMovePrompt, sourceNote.id, sourceQuote.id, destQuote.id)
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Library
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Quote, alter its Quote Number
                    newNote.quote_num = destQuote.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Quote Number
                        sourceNote.quote_num = destQuote.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Quote "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destQuote.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.
                            nodeList = (_('Collections'),) + sourceQuote.GetNodeData() + (sourceNote.id, )
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'QuoteNoteNode')
                                 
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Quote "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destQuote.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Note Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Collections'),) + destQuote.GetNodeData() + (tempNote.id,)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('QuoteNoteNode', nodeList, tempNote.number, tempNote.collection_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Quote Note message
                        msg = "AQN Collections >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                             msg += " >|< %s"
                             data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop a Clip Note on a new Clip
    elif (sourceData.nodetype == 'ClipNoteNode' and destNodeData.nodetype == 'ClipNode'):
        # Load the Source Clip Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Clip
        sourceClip = Clip.Clip(sourceNote.clip_num)
        # Load the Destination Clip
        destClip = Clip.Clip(destNodeData.recNum)
        # Can't drop a note on it's own parent!
        if sourceClip.number != destClip.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nClip "%s" to\nClip "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceClip.id, destClip.id)
            else:
                prompt = _('Do you want to %s Note "%s" from\nClip "%s" to\nClip "%s"?') % (copyMovePrompt, sourceNote.id, sourceClip.id, destClip.id)
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Library
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Clip, alter its Clip Number
                    newNote.clip_num = destClip.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Clip Number
                        sourceNote.clip_num = destClip.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Clip "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destClip.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.
                            nodeList = (_('Collections'),) + sourceClip.GetNodeData() + (sourceNote.id, )
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'ClipNoteNode')
                                 
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Clip "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destClip.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Note Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Collections'),) + destClip.GetNodeData() + (tempNote.id,)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('ClipNoteNode', nodeList, tempNote.number, tempNote.collection_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Clip Note message
                        msg = "AClN Collections >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                             msg += " >|< %s"
                             data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop a Snapshot Note on a new Snapshot
    elif (sourceData.nodetype == 'SnapshotNoteNode' and destNodeData.nodetype == 'SnapshotNode'):
        # Load the Source Snapshot Note
        sourceNote = Note.Note(id_or_num=sourceData.recNum)
        # Load the Source Snapshot
        sourceSnapshot = Snapshot.Snapshot(sourceNote.snapshot_num)
        # Load the Destination Snapshot
        destSnapshot = Snapshot.Snapshot(destNodeData.recNum)
        # Can't drop a note on it's own parent!
        if sourceSnapshot.number != destSnapshot.number:
            # Get user confirmation of the Note Copy/Move request
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Note "%s" from\nSnapshot "%s" to\nSnapshot "%s"?'), 'utf8') % (copyMovePrompt, sourceNote.id, sourceSnapshot.id, destSnapshot.id)
            else:
                prompt = _('Do you want to %s Note "%s" from\nSnapshot "%s" to\nSnapshot "%s"?') % (copyMovePrompt, sourceNote.id, sourceSnapshot.id, destSnapshot.id)
            # Display the prompt and get user input
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            dlg.Destroy()
            # If the user says YES ...
            if result == wx.ID_YES:
                # Copy or Move the Note to the Destination Library
                contin = True
                # If copying ...
                if action == 'Copy':
                    # Make a duplicate of the note to be copied
                    newNote = sourceNote.duplicate()
                    # To place the copy in the destination Snapshot, alter its Snapshot Number
                    newNote.snapshot_num = destSnapshot.number
                # If moving ...
                elif action == 'Move':
                    # We need to trap Record Lock exceptions
                    try:
                        # Lock the Note Record to prevent other users from altering it simultaneously
                        sourceNote.lock_record()
                        # To move a Note, alter its Snapshot Number
                        sourceNote.snapshot_num = destSnapshot.number
                    # If the record IS locked ...
                    except TransanaExceptions.RecordLockedError, e:
                        # Prepare the error message
                        if 'unicode' in wx.PlatformInfo:
                            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                            prompt = unicode(_('You cannot move Note "%s"') + \
                                             _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
                        else:
                            prompt = _('You cannot move Note "%s"') + \
                                     _('.\nThe record is currently locked by %s.\nPlease try again later.')
                        # Display the error message
                        errordlg = Dialogs.ErrorDialog(None, prompt % (sourceNote.id, e.user))
                        errordlg.ShowModal()
                        errordlg.Destroy()
                        # If an error arises, we do NOT continue!
                        contin = False
                # If no error has occurred yet ...
                if contin:
                    # If copying ...
                    if action == 'Copy':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            newNote.db_save()
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Snapshot "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destSnapshot.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                    # If moving ...
                    elif action == 'Move':
                        # Begin exception handling to trap save errors
                        try:
                            # Save the new Note to the database.
                            sourceNote.db_save()
                            # Remove the old Note from the Tree.
                            # delete_Node needs to be able to climb the tree, so we need to build the Node List that
                            # tells it what to delete.
                            nodeList = (_('Collections'),) + sourceSnapshot.GetNodeData() + (sourceNote.id, )
                            # Now request that the defined node be deleted.  (DeleteNode sends its own MU messages!
                            treeCtrl.delete_Node(nodeList, 'SnapshotNoteNode')
                                 
                        # If the Save fails ...
                        except TransanaExceptions.SaveError, e:
                            # Prepare the Error Message
                            msg = _('A Note named "%s" already exists in Snapshot "%s".')
                            if 'unicode' in wx.PlatformInfo:
                                msg = unicode(msg, 'utf8')
                            # Display the error message
                            errordlg = Dialogs.ErrorDialog(None, msg % (sourceNote.id, destSnapshot.id))
                            errordlg.ShowModal()
                            errordlg.Destroy()
                            # If an error arises, we do NOT continue!
                            contin = False
                        # Unlock the Note Record
                        sourceNote.unlock_record()

                        # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                        ClearClipboard()

                # If no error has occurred yet ...
                if contin:
                    # We need to update the database tree and inform other copies of MU.
                    # Start by getting the correct note object.
                    if action == 'Copy':
                        tempNote = newNote
                    elif action == 'Move':
                        tempNote = sourceNote
                    # Build the Node List for the tree control
                    nodeList = (_('Collections'),) + destSnapshot.GetNodeData() + (tempNote.id,)
                    # Add the Node to the Tree
                    treeCtrl.add_Node('SnapshotNoteNode', nodeList, tempNote.number, tempNote.collection_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        # Prepare an Add Snapshot Note message
                        msg = "ASnN Collections >|< %s"
                        # Convert the Node List to the form needed for messaging
                        data = (nodeList[1],)
                        for nd in nodeList[2:]:
                             msg += " >|< %s"
                             data += (nd, )
                        # Send the message
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Drop a SearchCollection on a SearchCollection (Nest a SearchCollection in another SearchCollection)
    elif (sourceData.nodetype == 'SearchCollectionNode' and \
          (destNodeData.nodetype == 'SearchCollectionNode' or destNodeData.nodetype == 'SearchResultsNode')):
        # NOTE:  SearchCollections, SearchQuotes, and SearchClips don't exist in the database.  Therefore, to copy or move them,
        #        all we need to do is manipulate Database Tree Nodes

        # First, check that we're not dragging a collection into a collection that's nested inside it.
        # Determine the parent node
        parentNode = treeCtrl.GetItemParent(destNode)
        # Keep climbing the node tree while we have valid nodes
        while parentNode.IsOk():
            # Get the parent node's data
            parentData = treeCtrl.GetPyData(parentNode)
            # if the parent node matches the source node, we have an error condition
            if (parentData.recNum == sourceData.recNum) and (parentData.nodetype == sourceData.nodetype):
                # Create a prompt for the user
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_("You cannot nest a Collection in one of its nested collections."), 'utf8')
                else:
                    prompt = _("You cannot nest a Collection in one of its nested collections.")
                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, prompt)
                errordlg.ShowModal()
                errordlg.Destroy()
                # We need to exit this method.
                return
            # Get the next parent node, one level up
            parentNode = treeCtrl.GetItemParent(parentNode)

        # Get user confirmation of the Collection Copy request.
        # First, set the names for use in the prompt.
        sourceCollectionId = sourceData.text
        destCollectionId = treeCtrl.GetItemText(destNode)
        # If we're dropping on the Search Results Node ...
        if destNodeData.nodetype == 'SearchResultsNode':
            # Create a prompt for the user
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Search Results Collection "%s" \nto the main Search Results node "%s"?'), 'utf8') % (copyMovePrompt, sourceCollectionId, destCollectionId)
            else:
                prompt = _('Do you want to %s Search Results Collection "%s" \nto the main Search Results node "%s"?') % (copyMovePrompt, sourceCollectionId, destCollectionId)
        # If we're dropping on a SearchCollectionNode ...
        else:
            # Create a prompt for the user
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Search Results Collection "%s" \nso it is nested in Search Results Collection "%s"?'), 'utf8') % (copyMovePrompt, sourceCollectionId, destCollectionId)
            else:
                prompt = _('Do you want to %s Search Results Collection "%s" \nso it is nested in Search Results Collection "%s"?') % (copyMovePrompt, sourceCollectionId, destCollectionId)
        # Get user confirmation of the Collection Copy request
        dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
        result = dlg.LocalShowModal()
        dlg.Destroy()
        # Process the results only if the user confirmed.
        if result == wx.ID_YES:
            # Let's get the Source SearchCollection.  However, we don't have direct access to the Source
            # Node in all cases.  (We can get it in Drag and Drop, but not Cut and Paste.  I tried to pass the actual
            # node in the SourceData, but that caused catastrophic program crashes.)
            # This ALWAYS gets the correct Source Node.
            sourceNode = treeCtrl.select_Node(sourceData.nodeList, sourceData.nodetype)

            # Let's make sure there isn't already a Search Collection node with the SourceNode's name in the DestNode
            # Get the Dest Collection's first Child Node
            (childNode, cookieItem) = treeCtrl.GetFirstChild(destNode)
            # Iterate through all the Source Collection's Child Nodes
            while childNode.IsOk():
                # Get the child node's data
                childData = treeCtrl.GetPyData(childNode)
                # If the nodetypes and item text match, we need an error message ...
                if (childData.nodetype == sourceData.nodetype) and \
                   (treeCtrl.GetItemText(childNode) == treeCtrl.GetItemText(sourceNode)):
                    # Create a prompt for the user
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('A Collection named "%s" already exists.'), 'utf8') % treeCtrl.GetItemText(childNode)
                    else:
                        prompt = _('A Collection named "%s" already exists.') % treeCtrl.GetItemText(childNode)
                    # Display the Error Message
                    errordlg = Dialogs.ErrorDialog(None, prompt)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # We need to exit this method.
                    return
                # If we are not currently looking at the Dest Collection's LAST Child ...
                if childNode != treeCtrl.GetLastChild(destNode):
                    # ... then let's move on to the next Child record.
                    (childNode, cookieItem) = treeCtrl.GetNextChild(destNode, cookieItem)
                # Otherwise ...
                else:
                    # We've seen all there is to see, so stop iterating.
                    break

            # Build the node list for the destination collection!
            # Start with the Destination Node name and build the node list up from the back
            destNodeList = (treeCtrl.GetItemText(destNode), )
            # Get the Destination Node's parent, and its node data
            parentNode = treeCtrl.GetItemParent(destNode)
            parentData = treeCtrl.GetPyData(parentNode)
            # As long as we get valid parent nodes, up to the point where we reach the Search Root Node ...
            while parentNode.IsOk() and \
                  (parentData.nodetype != "SearchRootNode"):
                # ... prepend the parent node's name to the node list ...
                destNodeList = (treeCtrl.GetItemText(parentNode),) + destNodeList
                # Get the Parent Node's parent, and its node data
                parentNode = treeCtrl.GetItemParent(parentNode)
                parentData = treeCtrl.GetPyData(parentNode)
            # Finally, add the Search Root to the node list
            destNodeList = (_('Search'),) + destNodeList

            # Now Add the new Node, using the SourceData's Data
            # No need to communicate with other Transana Clients here, we're just manipulating Search Results.

            # If we're MOVING ...
            if action == 'Move':
                # .. then MOVE the node
                treeCtrl.copy_Node(sourceData.nodetype, sourceData.nodeList, destNodeList, True, sendMessage=False)
                # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                ClearClipboard()
            # If we're COPYING ...
            else:
                # ... then COPY the node
                treeCtrl.copy_Node(sourceData.nodetype, sourceData.nodeList, destNodeList, False, sendMessage=False)
                
            # Select the Destination Collection as the tree's Selected Item
            treeCtrl.UnselectAll()
            treeCtrl.SelectItem(destNode)

    # Drop a SearchQuote on a SearchCollection (Copy or Move a SearchQuote)
    elif (sourceData.nodetype == 'SearchQuoteNode' and destNodeData.nodetype == 'SearchCollectionNode'):
        # NOTE:  SearchQuotes don't exist in the database.  Therefore, to copy or move them,
        #        all we need to do is manipulate Database Tree Nodes

        # Get user confirmation of the Quote Copy/Move request.
        # First, let's get the appropriate text for the confirmation prompt.
        sourceQuoteId = sourceData.text
        # The Source Collection is the second-to-last entry in the source Node List!
        sourceCollectionId = sourceData.nodeList[-2]
        destCollectionId = treeCtrl.GetItemText(destNode)
        # Create the confirmation Dialog box
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to %s Search Results Quote "%s" from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?'), 'utf8') % (copyMovePrompt, sourceQuoteId, sourceCollectionId, destCollectionId)
        else:
            prompt = _('Do you want to %s Search Results Quote "%s" from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?') % (copyMovePrompt, sourceQuoteId, sourceCollectionId, destCollectionId)
        # Display the confirmation Dialog Box
        dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
        result = dlg.LocalShowModal()
        # Clean up after the confirmation Dialog box
        dlg.Destroy()
        # If the user confirmed, process the request.
        if result == wx.ID_YES:
            # Check for Duplicate Quote Name error
            (dupResult, newQuoteName) = CheckForDuplicateObjName(sourceData.text, 'SearchQuoteNode', treeCtrl, destNode)
            # If a Duplicate Quote Name exists that is not resolved within CheckForDuplicateObjNames ...
            if dupResult:
                # ... display the error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('%s cancelled for Quote "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
                else:
                    prompt = _('%s cancelled for Quote "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceData.text)
                dlg = Dialogs.ErrorDialog(None, prompt)
                dlg.ShowModal()
                dlg.Destroy()
            else:
                # The user may have provided a new name for the Quote if it was a duplicate.
                if newQuoteName != sourceData.text:
                    # If so, use this new name.
                    sourceData.text = newQuoteName

                # What we need to do first is add the appropriate new Node to the Tree.
                # Let's build the NodeList by starting with the Source Quote text ...
                nodeList = (sourceData.text,)
                # ... and then climbing the Destination Node Tree .... 
                currentNode = destNode
                # ... until we get to the SearchRootNode
                while treeCtrl.GetPyData(currentNode).nodetype != 'SearchRootNode':
                    # We add the Item Test to the FRONT of the Node List
                    nodeList = (treeCtrl.GetItemText(currentNode),) + nodeList
                    # and we move up to the node's parent
                    currentNode = treeCtrl.GetItemParent(currentNode)
                # Get the node data for the DESTINATION node, so we can update the Quote's Parent record
                destData = treeCtrl.GetPyData(destNode)
                # Now Add the new Node, using the SourceData's Data
                treeCtrl.add_Node('SearchQuoteNode', (_('Search'),) + nodeList, sourceData.recNum, destData.recNum, expandNode=False)
                # No need to communicate with other Transana Clients here, we're just manipulating Search Results.
                # If we need to remove the node, the SourceData carries the nodeList we need to delete
                if action == 'Move':
                    treeCtrl.delete_Node(sourceData.nodeList, 'SearchQuoteNode')
                    # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                    ClearClipboard()
                # Select the Destination Collection as the tree's Selected Item
                treeCtrl.SelectItem(destNode)

    # Drop a SearchQuote on a SearchQuote, SearchClip or SearchSnapshot (Copy or Move a SearchQuote into a position in the SortOrder)
    elif (sourceData.nodetype == 'SearchQuoteNode' and destNodeData.nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']):
        # NOTE:  SearchQuotes don't exist in the database.  Therefore, to copy or move them,
        #        all we need to do is manipulate Database Tree Nodes
        # Set variables for the user Confirmation Prompt, if needed.
        sourceQuoteId = sourceData.text
        # The Source Collection is the second-to-last entry in the source Node List!
        sourceCollectionId = sourceData.nodeList[-2]
        destCollectionId = treeCtrl.GetItemText(treeCtrl.GetItemParent(destNode))
        # Only prompt if we are copying/moving to a DIFFERENT collection, not the same one!
        if (sourceCollectionId == destCollectionId) and \
            (sourceData.parent == destNodeData.parent):
            # This bypasses the prompt if both collection names and parents are the same
            result = wx.ID_YES
            # By definition, this MUST be a move, as we are altering Sort Order.
            action = 'Move'
            # We don't have to deal with duplicate Clip Names if changing a Clip's Sort Order.
            # It IS a duplicate name, but we can ignore that.
            dupResult = False
        else:
            # If we're dealing with different collections, prompt the user to confirm the copy/move.
            # First, build the confirmation Dialog Box
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Quote "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceQuoteId, sourceCollectionId, destCollectionId)
            else:
                prompt = _('Do you want to %s Quote "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceQuoteId, sourceCollectionId, destCollectionId)
            # Display the Confirmation Dialog Box
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            # Clean up the Confirmation Dialog Box
            dlg.Destroy()
            # If the user confirmed ...
            if result == wx.ID_YES:
                # ... we need to check to see if this Quote is a Duplicate, giving the user a chance to resolve the problem.
                (dupResult, newQuoteName) = CheckForDuplicateObjName(sourceData.text, 'SearchQuoteNode', treeCtrl, treeCtrl.GetItemParent(destNode))
                # If the quote is no longer a duplicate and the user has changed the quote's name to resolve that ...
                if (not dupResult) and (newQuoteName != sourceData.text):
                   # ... use the new name.
                   sourceData.text = newQuoteName

        # If the user confirmed (or wasn't asked) ...
        if result == wx.ID_YES:
            # If we have a Duplicate Quote Name error ...
            if dupResult:
                # ... show the error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('%s cancelled for Quote "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
                else:
                    prompt = _('%s cancelled for Quote "%s".  Duplicate Quote Name Error.') % (copyMovePrompt, sourceData.text)
                dlg = Dialogs.ErrorDialog(None, prompt)
                dlg.ShowModal()
                dlg.Destroy()
            # If there's no Duplicate Clip Name Error ...
            else:
                # What we need to do first is add the appropriate new Node to the Tree.
                # Let's build the NodeList by starting with the Source Quote text ...
                nodeList = (sourceData.text,)
                # ... and then climbing the Destination Node Tree  using the quote's Parent Collection .... 
                currentNode = treeCtrl.GetItemParent(destNode)
                # ... until we get to the SearchRootNode
                while treeCtrl.GetPyData(currentNode).nodetype != 'SearchRootNode':
                    # We add the Item Test to the FRONT of the Node List
                    nodeList = (treeCtrl.GetItemText(currentNode),) + nodeList
                    # and we move up to the node's parent
                    currentNode = treeCtrl.GetItemParent(currentNode)

                # If we need to remove the node, the SourceData carries the nodeList we need to delete.
                # (We need to delete first so that the moved Quote in the same Collection doesn't get removed
                # immediately after being added.)
                if action == 'Move':
                    treeCtrl.delete_Node(sourceData.nodeList, sourceData.nodetype)
                    # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                    ClearClipboard()
                # Now Add the new Node, using the SourceData's Data but the destNode's Parent
                treeCtrl.add_Node('SearchQuoteNode', (_('Search'),) + nodeList, sourceData.recNum, destNodeData.parent, expandNode=False, insertPos=destNode)
                # No need to communicate with other Transana Clients here, we're just manipulating Search Results.
                # Select the Destination Collection as the tree's Selected Item
                treeCtrl.SelectItem(destNode)

    # Drop a SearchClip on a SearchCollection (Copy or Move a SearchClip)
    elif (sourceData.nodetype == 'SearchClipNode' and destNodeData.nodetype == 'SearchCollectionNode'):
        # NOTE:  SearchClips don't exist in the database.  Therefore, to copy or move them,
        #        all we need to do is manipulate Database Tree Nodes

        # Get user confirmation of the Clip Copy/Move request.
        # First, let's get the appropriate text for the confirmation prompt.
        sourceClipId = sourceData.text
        # The Source Collection is the second-to-last entry in the source Node List!
        sourceCollectionId = sourceData.nodeList[-2]
        destCollectionId = treeCtrl.GetItemText(destNode)
        # Create the confirmation Dialog box
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to %s Search Results Clip "%s" from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?'), 'utf8') % (copyMovePrompt, sourceClipId, sourceCollectionId, destCollectionId)
        else:
            prompt = _('Do you want to %s Search Results Clip "%s" from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?') % (copyMovePrompt, sourceClipId, sourceCollectionId, destCollectionId)
        # Display the confirmation Dialog Box
        dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
        result = dlg.LocalShowModal()
        # Clean up after the confirmation Dialog box
        dlg.Destroy()
        # If the user confirmed, process the request.
        if result == wx.ID_YES:
            # Check for Duplicate Clip Name error
            (dupResult, newClipName) = CheckForDuplicateObjName(sourceData.text, 'SearchClipNode', treeCtrl, destNode)
            # If a Duplicate Clip Name exists that is not resolved within CheckForDuplicateObjNames ...
            if dupResult:
                # ... display the error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('%s cancelled for Clip "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
                else:
                    prompt = _('%s cancelled for Clip "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceData.text)
                dlg = Dialogs.ErrorDialog(None, prompt)
                dlg.ShowModal()
                dlg.Destroy()
            else:
                # The user may have provided a new name for the Clip if it was a duplicate.
                if newClipName != sourceData.text:
                    # If so, use this new name.
                    sourceData.text = newClipName

                # What we need to do first is add the appropriate new Node to the Tree.
                # Let's build the NodeList by starting with the Source Clip text ...
                nodeList = (sourceData.text,)
                # ... and then climbing the Destination Node Tree .... 
                currentNode = destNode
                # ... until we get to the SearchRootNode
                while treeCtrl.GetPyData(currentNode).nodetype != 'SearchRootNode':
                    # We add the Item Test to the FRONT of the Node List
                    nodeList = (treeCtrl.GetItemText(currentNode),) + nodeList
                    # and we move up to the node's parent
                    currentNode = treeCtrl.GetItemParent(currentNode)
                # Get the node data for the DESTINATION node, so we can update the Clip's Parent record
                destData = treeCtrl.GetPyData(destNode)
                # Now Add the new Node, using the SourceData's Data
                treeCtrl.add_Node('SearchClipNode', (_('Search'),) + nodeList, sourceData.recNum, destData.recNum, expandNode=False)
                # No need to communicate with other Transana Clients here, we're just manipulating Search Results.
                # If we need to remove the node, the SourceData carries the nodeList we need to delete
                if action == 'Move':
                    treeCtrl.delete_Node(sourceData.nodeList, 'SearchClipNode')
                    # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                    ClearClipboard()
                # Select the Destination Collection as the tree's Selected Item
                treeCtrl.SelectItem(destNode)

    # Drop a SearchClip on a SearchQuote, SearchClip or SearchSnapshot (Copy or Move a SearchClip into a position in the SortOrder)
    elif (sourceData.nodetype == 'SearchClipNode' and destNodeData.nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']):
        # NOTE:  SearchClips don't exist in the database.  Therefore, to copy or move them,
        #        all we need to do is manipulate Database Tree Nodes
        # Set variables for the user Confirmation Prompt, if needed.
        sourceClipId = sourceData.text
        # The Source Collection is the second-to-last entry in the source Node List!
        sourceCollectionId = sourceData.nodeList[-2]
        destCollectionId = treeCtrl.GetItemText(treeCtrl.GetItemParent(destNode))
        # Only prompt if we are copying/moving to a DIFFERENT collection, not the same one!
        if (sourceCollectionId == destCollectionId) and \
            (sourceData.parent == destNodeData.parent):
            # This bypasses the prompt if both collection names and parents are the same
            result = wx.ID_YES
            # By definition, this MUST be a move, as we are altering Sort Order.
            action = 'Move'
            # We don't have to deal with duplicate Clip Names if changing a Clip's Sort Order.
            # It IS a duplicate name, but we can ignore that.
            dupResult = False
        else:
            # If we're dealing with different collections, prompt the user to confirm the copy/move.
            # First, build the confirmation Dialog Box
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceClipId, sourceCollectionId, destCollectionId)
            else:
                prompt = _('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceClipId, sourceCollectionId, destCollectionId)
            # Display the Confirmation Dialog Box
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            # Clean up the Confirmation Dialog Box
            dlg.Destroy()
            # If the user confirmed ...
            if result == wx.ID_YES:
                # ... we need to check to see if this Clip is a Duplicate, giving the user a chance to resolve the problem.
                (dupResult, newClipName) = CheckForDuplicateObjName(sourceData.text, 'SearchClipNode', treeCtrl, treeCtrl.GetItemParent(destNode))
                # If the clip is no longer a duplicate and the user has changed the clip's name to resolve that ...
                if (not dupResult) and (newClipName != sourceData.text):
                   # ... use the new name.
                   sourceData.text = newClipName

        # If the user confirmed (or wasn't asked) ...
        if result == wx.ID_YES:
            # If we have a Duplicate Clip Name error ...
            if dupResult:
                # ... show the error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('%s cancelled for Clip "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
                else:
                    prompt = _('%s cancelled for Clip "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceData.text)
                dlg = Dialogs.ErrorDialog(None, prompt)
                dlg.ShowModal()
                dlg.Destroy()
            # If there's no Duplicate Clip Name Error ...
            else:
                # What we need to do first is add the appropriate new Node to the Tree.
                # Let's build the NodeList by starting with the Source Clip text ...
                nodeList = (sourceData.text,)
                # ... and then climbing the Destination Node Tree  using the clip's Parent Collection .... 
                currentNode = treeCtrl.GetItemParent(destNode)
                # ... until we get to the SearchRootNode
                while treeCtrl.GetPyData(currentNode).nodetype != 'SearchRootNode':
                    # We add the Item Test to the FRONT of the Node List
                    nodeList = (treeCtrl.GetItemText(currentNode),) + nodeList
                    # and we move up to the node's parent
                    currentNode = treeCtrl.GetItemParent(currentNode)

                # If we need to remove the node, the SourceData carries the nodeList we need to delete.
                # (We need to delete first so that the moved Clip in the same Collection doesn't get removed
                # immediately after being added.)
                if action == 'Move':
                    treeCtrl.delete_Node(sourceData.nodeList, sourceData.nodetype)
                    # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                    ClearClipboard()
                # Now Add the new Node, using the SourceData's Data but the destNode's Parent
                treeCtrl.add_Node('SearchClipNode', (_('Search'),) + nodeList, sourceData.recNum, destNodeData.parent, expandNode=False, insertPos=destNode)
                # No need to communicate with other Transana Clients here, we're just manipulating Search Results.
                # Select the Destination Collection as the tree's Selected Item
                treeCtrl.SelectItem(destNode)

    # Drop a SearchSnapshot on a SearchCollection (Copy or Move a SearchSnapshot)
    elif (sourceData.nodetype == 'SearchSnapshotNode' and destNodeData.nodetype == 'SearchCollectionNode'):
        # NOTE:  SearchSnapshot don't exist in the database.  Therefore, to copy or move them,
        #        all we need to do is manipulate Database Tree Nodes

        # Get user confirmation of the Snapshot Copy/Move request.
        # First, let's get the appropriate text for the confirmation prompt.
        sourceSnapshotId = sourceData.text
        # The Source Collection is the second-to-last entry in the source Node List!
        sourceCollectionId = sourceData.nodeList[-2]
        destCollectionId = treeCtrl.GetItemText(destNode)
        # Create the confirmation Dialog box
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to %s Search Results Snapshot "%s" from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?'), 'utf8') % (copyMovePrompt, sourceSnapshotId, sourceCollectionId, destCollectionId)
        else:
            prompt = _('Do you want to %s Search Results Snapshot "%s" from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?') % (copyMovePrompt, sourceSnapshotId, sourceCollectionId, destCollectionId)
        # Display the confirmation Dialog Box
        dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
        result = dlg.LocalShowModal()
        # Clean up after the confirmation Dialog box
        dlg.Destroy()
        # If the user confirmed, process the request.
        if result == wx.ID_YES:
            # Check for Duplicate Snapshot Name error
            (dupResult, newSnapshotName) = CheckForDuplicateObjName(sourceData.text, 'SearchSnapshotNode', treeCtrl, destNode)
            # If a Duplicate Snapshot Name exists that is not resolved within CheckForDuplicateObjNames ...
            if dupResult:
                # ... display the error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('%s cancelled for Snapshot "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
                else:
                    prompt = _('%s cancelled for Snapshot "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceData.text)
                dlg = Dialogs.ErrorDialog(None, prompt)
                dlg.ShowModal()
                dlg.Destroy()
            else:
                # The user may have provided a new name for the Clip if it was a duplicate.
                if newSnapshotName != sourceData.text:
                    # If so, use this new name.
                    sourceData.text = newSnapshotName

                # What we need to do first is add the appropriate new Node to the Tree.
                # Let's build the NodeList by starting with the Source Clip text ...
                nodeList = (sourceData.text,)
                # ... and then climbing the Destination Node Tree .... 
                currentNode = destNode
                # ... until we get to the SearchRootNode
                while treeCtrl.GetPyData(currentNode).nodetype != 'SearchRootNode':
                    # We add the Item Test to the FRONT of the Node List
                    nodeList = (treeCtrl.GetItemText(currentNode),) + nodeList
                    # and we move up to the node's parent
                    currentNode = treeCtrl.GetItemParent(currentNode)
                # Get the node data for the DESTINATION node, so we can update the Clip's Parent record
                destData = treeCtrl.GetPyData(destNode)
                # Now Add the new Node, using the SourceData's Data
                treeCtrl.add_Node('SearchSnapshotNode', (_('Search'),) + nodeList, sourceData.recNum, destData.recNum, expandNode=False)
                # No need to communicate with other Transana Clients here, we're just manipulating Search Results.
                # If we need to remove the node, the SourceData carries the nodeList we need to delete
                if action == 'Move':
                    treeCtrl.delete_Node(sourceData.nodeList, 'SearchSnapshotNode')
                    # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                    ClearClipboard()
                # Select the Destination Collection as the tree's Selected Item
                treeCtrl.SelectItem(destNode)

    # Drop a SearchSnapshot on a SearchQuote, SearchClip or SearchSnapshot (Copy or Move a SearchSnapshot into a position in the SortOrder)
    elif (sourceData.nodetype == 'SearchSnapshotNode' and destNodeData.nodetype in ['SearchQuoteNode', 'SearchClipNode', 'SearchSnapshotNode']):
        # NOTE:  SearchSnapshots don't exist in the database.  Therefore, to copy or move them,
        #        all we need to do is manipulate Database Tree Nodes

        # Set variables for the user Confirmation Prompt, if needed.
        sourceSnapshotId = sourceData.text
        # The Source Collection is the second-to-last entry in the source Node List!
        sourceCollectionId = sourceData.nodeList[-2]
        destCollectionId = treeCtrl.GetItemText(treeCtrl.GetItemParent(destNode))

        # Only prompt if we are copying/moving to a DIFFERENT collection, not the same one!
        if (sourceCollectionId == destCollectionId) and \
            (sourceData.parent == destNodeData.parent):
            # This bypasses the prompt if both collection names and parents are the same
            result = wx.ID_YES
            # By definition, this MUST be a move, as we are altering Sort Order.
            action = 'Move'
            # We don't have to deal with duplicate Clip Names if changing a Clip's Sort Order.
            # It IS a duplicate name, but we can ignore that.
            dupResult = False
        else:
            # If we're dealing with different collections, prompt the user to confirm the copy/move.
            # First, build the confirmation Dialog Box
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('Do you want to %s Snapshot "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceSnapshotId, sourceCollectionId, destCollectionId)
            else:
                prompt = _('Do you want to %s Snapshot "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceSnapshotId, sourceCollectionId, destCollectionId)
            # Display the Confirmation Dialog Box
            dlg = Dialogs.QuestionDialog(treeCtrl, prompt)
            result = dlg.LocalShowModal()
            # Clean up the Confirmation Dialog Box
            dlg.Destroy()
            # If the user confirmed ...
            if result == wx.ID_YES:
                # ... we need to check to see if this Snapshot is a Duplicate, giving the user a chance to resolve the problem.
                (dupResult, newSnapshotName) = CheckForDuplicateObjName(sourceData.text, 'SearchSnapshotNode', treeCtrl, treeCtrl.GetItemParent(destNode))
                # If the clip is no longer a duplicate and the user has changed the clip's name to resolve that ...
                if (not dupResult) and (newSnapshotName != sourceData.text):
                   # ... use the new name.
                   sourceData.text = newSnapshotName

        # If the user confirmed (or wasn't asked) ...
        if result == wx.ID_YES:
            # If we have a Duplicate Snapshot Name error ...
            if dupResult:
                # ... show the error message.
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('%s cancelled for Snapshot "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
                else:
                    prompt = _('%s cancelled for Snapshot "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceData.text)
                dlg = Dialogs.ErrorDialog(None, prompt)
                dlg.ShowModal()
                dlg.Destroy()
            # If there's no Duplicate Snapshot Name Error ...
            else:
                # What we need to do first is add the appropriate new Node to the Tree.
                # Let's build the NodeList by starting with the Source Clip text ...
                nodeList = (sourceData.text,)
                # ... and then climbing the Destination Node Tree  using the clip's Parent Collection .... 
                currentNode = treeCtrl.GetItemParent(destNode)
                # ... until we get to the SearchRootNode
                while treeCtrl.GetPyData(currentNode).nodetype != 'SearchRootNode':
                    # We add the Item Test to the FRONT of the Node List
                    nodeList = (treeCtrl.GetItemText(currentNode),) + nodeList
                    # and we move up to the node's parent
                    currentNode = treeCtrl.GetItemParent(currentNode)

                # If we need to remove the node, the SourceData carries the nodeList we need to delete.
                # (We need to delete first so that the moved Clip in the same Collection doesn't get removed
                # immediately after being added.)
                if action == 'Move':
                    treeCtrl.delete_Node(sourceData.nodeList, sourceData.nodetype)
                    # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
                    ClearClipboard()
                # Now Add the new Node, using the SourceData's Data, but the destination node's parent
                treeCtrl.add_Node('SearchSnapshotNode', (_('Search'),) + nodeList, sourceData.recNum, destNodeData.parent, expandNode=False, insertPos=destNode)
                # No need to communicate with other Transana Clients here, we're just manipulating Search Results.
                # Select the Destination Collection as the tree's Selected Item
                treeCtrl.SelectItem(destNode)

    else:
        # This code will only get called if a source/drop combination is defined in DragDropEvaluation()
        # as a legal drop but is not defined here with the appropriate code to define what to do!
        # Notify the user of the problem.  This should never be seen by the general public, but is intended for developers.
        dlg = Dialogs.ErrorDialog(treeCtrl, 'Drop of %s on %s allowed by "DragDropEvaluation()" but not implemented in "ProcessPasteDrop()".' % (sourceData.nodetype, destNodeData.nodetype))
        dlg.ShowModal()
        dlg.Destroy()

def ClearClipboard():
    """ If we Moved, we need to clear out the Clipboard.  We'll do this by putting an empty DataTreeDragDropData object into the Clipboard. """
    # Create an empty DataTreeDragDropData object
    ddd = DataTreeDragDropData()
    # Use cPickle to convert the data object into a string representation
    pddd = cPickle.dumps(ddd, 1)
    # Now create a wxCustomDataObject for dragging and dropping and
    # assign it a custom Data Format
    cdo = wx.CustomDataObject(wx.CustomDataFormat('DataTreeDragData'))
    # Put the pickled data object in the wxCustomDataObject
    cdo.SetData(pddd)
    # Is the clipboard open?  Assume not.
    clipboardOpenedHere = False
    # If the Clipboard is NOT open ...
    if not wx.TheClipboard.IsOpened():
        # ... open the Clipboard
        wx.TheClipboard.Open()
        # Remember that the Clipboard was opened HERE!
        clipboardOpenedHere = True
    # Now put the data in the clipboard.
    wx.TheClipboard.SetData(cdo)
    # If the clipboard was opened HERE ...
    if clipboardOpenedHere:
        # ... close the Clipboard
        wx.TheClipboard.Close()

def CheckForDuplicateObjName(sourceObjName, sourceObjType, treeCtrl, destCollectionNode):
   """ Check the destCollectionNode to see if sourceObjName already exists.  If so, prompt for a name change.
       Return True if duplicate is found, False if no duplicate is found or if the Object is renamed appropriately.  """
   # Before we do anything, let's make sure we have a Collection Node, not a Quote Node, Clip Node or Shapshot Node here
   if treeCtrl.GetPyData(destCollectionNode).nodetype in ['QuoteNode', 'ClipNode', 'SnapshotNode']:
       # If it's a quote, clip or snapshot, let's use its parent Collection
       destCollectionNode = treeCtrl.GetItemParent(destCollectionNode)
   # Assume that no duplicate exists unless proven otherwise
   result = False
   # Check to see if the Collection has children.  Otherwise, forget checking!
   if treeCtrl.ItemHasChildren(destCollectionNode):
      # wxTreeCtrl requires the "cookie" value to list children.  Initialize it.
      cookieVal = 0
      # Get the first child of the destCollectionNode 
      (tempTreeItem, cookieVal) = treeCtrl.GetFirstChild(destCollectionNode)
      # Iterate through all the destNode's children
      while tempTreeItem.IsOk():
         # Get the current child's Node Data
         tempData = treeCtrl.GetPyData(tempTreeItem)
         
         # See if the current child is a Quote or Clip AND it has the same name as the source Object
         if (tempData.nodetype in ['QuoteNode', 'ClipNode'])  and (treeCtrl.GetItemText(tempTreeItem).upper() == sourceObjName.upper()):
            # If so, prompt the user to change the object's Name.  First, build a Dialog to ask that question.
            dlg = wx.TextEntryDialog(TransanaGlobal.menuWindow, _('Duplicate Item Name.  Please enter a new name for the Item.'), _('Transana Error'), sourceObjName, style=wx.OK | wx.CANCEL | wx.CENTRE)
            # Position the Dialog Box in the center of the screen
            dlg.CentreOnScreen()
            # Show the Dialog Box
            dlgResult = dlg.ShowModal()
            # If the user selected OK AND changed the Item Name ...
            if (dlgResult == wx.ID_OK) and (dlg.GetValue() != sourceObjName):
               # Let's look at the new name ...
               sourceObjName = dlg.GetValue()
               # ... and see if it is a Duplicate Object Name by recursively calling this method
               (result, sourceObjName) = CheckForDuplicateObjName(sourceObjName, sourceObjType, treeCtrl, destCollectionNode)
               # Clean up after prompting for the new Object name
               dlg.Destroy()
               # We can stop looking for duplicate names.  We've already found it.
               break
            else:
               # If the user selected CANCEL or failed to change the Object Name,
               # we set result to True to indicate that a Duplicate Object Name was found.
               result = True
               # Clean up after prompting for a new Object Name
               dlg.Destroy()
               # We can stop looking for duplicate names.  We've already found it.
               break

         # If we're not at the last child, get the next child for dropNode
         if tempTreeItem != treeCtrl.GetLastChild(destCollectionNode):
            (tempTreeItem, cookieVal) = treeCtrl.GetNextChild(destCollectionNode, cookieVal)
         # If we are at the last child, exit the while loop.
         else:
            break

   # Return the result and the Object Name, as it could have been changed.
   return (result, sourceObjName)

def CopyMoveQuote(treeCtrl, destNode, sourceQuote, sourceCollection, destCollection, action):
    """ This function copies or moves sourceQuote to destCollection, depending on the value of 'action' """
    contin = True
    if action == 'Copy':
       # Make a duplicate of the quote to be copied
       newQuote = sourceQuote.duplicate()
       # To place the copy in the destination collection, alter its Collection Number and Sort Order value
       newQuote.collection_num = destCollection.number
    elif action == 'Move':
        try:
           # Lock the Quote Record to prevent other users from altering it simultaneously
           sourceQuote.lock_record()
           # To move a quote, alter its Collection Number and Sort Order value
           sourceQuote.collection_num = destCollection.number
        except TransanaExceptions.RecordLockedError, e:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('You cannot move Quote "%s"') + \
                                 _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
            else:
                prompt = _('You cannot move Quote "%s"') + \
                         _('.\nThe record is currently locked by %s.\nPlease try again later.')
            errordlg = Dialogs.ErrorDialog(None, prompt % (sourceQuote.id, e.user))
            errordlg.ShowModal()
            errordlg.Destroy()
            contin = False
    if contin:
       # NOTE:  CopyMoveQuote places the copy at the end of the Collection's Item List.  If that's not
       #        what we want, we can call ChangeClipOrder later.

       # Get the highest SortOrder value and add one to it 
       itemCount = DBInterface.getMaxSortOrder(destCollection.number) + 1

       # Check for Duplicate Clip Names, an error condition
       # First, get the name of the appropriate Quote Object
       if action == 'Copy':
          quoteName = newQuote.id
       elif action == 'Move':
          quoteName = sourceQuote.id
       # See if the Quote Name already exists in the Destination Collection
       (dupResult, newQuoteName) = CheckForDuplicateObjName(quoteName, 'QuoteNode', treeCtrl, destNode)
       
       # If a Duplicate Quote Name is found and the error situation not resolved, show an Error Message
       if dupResult:
          # Unlock the source quote (before presenting the dialog to keep if from being locked by a slow user response.)
          if action == 'Move':
              sourceQuote.unlock_record()
          # Report the failure to the user, although it's already known to have failed because they pressed "cancel".
          if 'unicode' in wx.PlatformInfo:
              # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
              prompt = unicode(_('%s cancelled for Quote "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceQuote.id)
          else:
              prompt = _('%s cancelled for Quote "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceQuote.id)
          dlg = Dialogs.ErrorDialog(treeCtrl, prompt)
          dlg.ShowModal()
          dlg.Destroy()
          return None
       else:
          # The user may have given the quote a new Quote Name in CheckForDuplicateObjName.  
          if newQuoteName != quoteName:
             # If so, use this new name!
             if action == 'Copy':
                newQuote.id = newQuoteName
             elif action == 'Move':
                 sourceQuote.id = newQuoteName 

          if action == 'Copy':
             # Now that we know the number of items in the collection, assign that as sortOrder
             newQuote.sort_order = itemCount
             # Save the new Quote to the database.
             newQuote.db_save()
          elif action == 'Move':
             # Now that we know the number of items in the collection, assign that as sortOrder
             sourceQuote.sort_order = itemCount
             # Save the new Quote to the database.
             sourceQuote.db_save()
             # Unlock the Quote Record
             sourceQuote.unlock_record()
                  
             # Remove the old Quote from the Tree.
             # delete_Node needs to be able to climb the tree, so we need to build the Node List that
             # tells it what to delete.  Start with the sourceCollection.
             nodeList = (sourceCollection.id,)
             # Load the Current Collection so we can find out about its parent, and work backwards from here.
             tempCollection = Collection.Collection(sourceCollection.number)
             # While the Current Collection has a defined parent...
             while tempCollection.parent > 0:
                # Make the Parent the Current Collection
                tempCollection = Collection.Collection(tempCollection.parent)
                # Add the Parent (now the Current Collection) to the FRONT of the Node List
                nodeList = (tempCollection.id,) + nodeList
             # Now add the Collections Root to the front of the Node List and the Quote's original name to the end of the Node List
             nodeList = (_('Collections'), ) + nodeList + (quoteName, )
             # Now request that the defined node be deleted
             treeCtrl.delete_Node(nodeList, 'QuoteNode')
             
             # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
             ClearClipboard()

          # Add the new Quote to the Database Tree Tab
          # To add a Quote, we need to build the node list for the tree's add_Node method to climb.
          # We need to add all of the Collection Parents to our Node List, so we'll start by loading
          # the current Collection
          tempCollection = Collection.Collection(destCollection.number)
          # Add the current Collection Name, and work backwards from here.
          nodeList = (tempCollection.id,)
          # Repeat this process as long as the Collection we're looking at has a defined Parent...
          while tempCollection.parent > 0:
             # Load the Parent Collection
             tempCollection = Collection.Collection(tempCollection.parent)
             # Add this Collection's name to the FRONT of the Node List
             nodeList = (tempCollection.id,) + nodeList
          # Now add the Collections Root node to the front of the Node List and the
          # Quote Name to the back of the Node List
          if action == 'Copy':
             nodeList = (_('Collections'), ) + nodeList + (newQuote.id, )
             # Add the Node to the Tree
             treeCtrl.add_Node('QuoteNode', nodeList, newQuote.number, newQuote.collection_num, sortOrder=newQuote.sort_order)

             # Now let's communicate with other Transana instances if we're in Multi-user mode
             if not TransanaConstants.singleUserVersion:
                msg = "AQ %s"
                data = (nodeList[1],)

                for nd in nodeList[2:]:
                   msg += " >|< %s"
                   data += (nd, )
                if TransanaGlobal.chatWindow != None:
                   TransanaGlobal.chatWindow.SendMessage(msg % data)

             # Now let's see if the Quote had Quote notes to copy!
             quoteNoteList = DBInterface.list_of_notes(Quote=sourceQuote.number, includeNumber=True)
             # For each note in the note list ...
             for quoteNote in quoteNoteList:
                 # Open a temporary copy of the note
                 tmpNote = Note.Note(quoteNote[0])
                 # Duplicate the note (which automatically gives it an object number = 0)
                 newNote = tmpNote.duplicate()
                 # Assign the duplicate to the NEW quote
                 newNote.quote_num = newQuote.number
                 # Save the new note
                 newNote.db_save()
                 # Add the new Note Node to the Tree
                 treeCtrl.add_Node('QuoteNoteNode', nodeList + (newNote.id,), newNote.number, newQuote.number)

                 # If the Notes Browser is open ...
                 if treeCtrl.parent.ControlObject.NotesBrowserWindow != None:
                     # ... Add the new note to the Notes Browser
                     treeCtrl.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('A', newNote)
                
                 # Now let's communicate with other Transana instances if we're in Multi-user mode
                 if not TransanaConstants.singleUserVersion:
                     # Prepare an Add Quote Note message
                     msg = "AQN Collections >|< %s"
                     # Convert the Node List to the form needed for messaging
                     data = (nodeList[1],)
                     for nd in (nodeList[2:] + (newNote.id,)):
                          msg += " >|< %s"
                          data += (nd, )
                     # Send the message
                     if TransanaGlobal.chatWindow != None:
                         TransanaGlobal.chatWindow.SendMessage(msg % data)

             # When copying a Quote and setting its sort order, we need to keep working with the new quote
             # rather than the old one.  Having CopyMoveQuote return the new quote makes this easy.
             return newQuote
          elif action == 'Move':
             nodeList = (_('Collections'), ) + nodeList + (sourceQuote.id, )
             # Add the Node to the Tree
             treeCtrl.add_Node('QuoteNode', nodeList, sourceQuote.number, sourceQuote.collection_num, sortOrder=sourceQuote.sort_order)
             # If we are moving a Quote, the Quote's Notes need to travel with the Quote.  The first step is to
             # get a list of those Notes.
             noteList = DBInterface.list_of_notes(Quote=sourceQuote.number)
             # If there are Quote Notes, we need to make sure they travel with the Quote
             if noteList != []:
                 newNode = treeCtrl.select_Node(nodeList, 'QuoteNode')
                 # We accomplish this using the TreeCtrl's "add_note_nodes" method
                 treeCtrl.add_note_nodes(noteList, newNode, Quote=sourceQuote.number)
                 treeCtrl.Refresh()

             # Now let's communicate with other Transana instances if we're in Multi-user mode
             if not TransanaConstants.singleUserVersion:
                msg = "AQ %s"
                data = (nodeList[1],)

                for nd in nodeList[2:]:
                   msg += " >|< %s"
                   data += (nd, )
                if TransanaGlobal.chatWindow != None:
                   TransanaGlobal.chatWindow.SendMessage(msg % data)

             # When copying a Quote and setting its sort order, we need to keep working with the new Quote
             # rather than the old one.  Having CopyMoveQuote return the new Quote makes this easy.
             return sourceQuote

def CopyMoveClip(treeCtrl, destNode, sourceClip, sourceCollection, destCollection, action):
    """ This function copies or moves sourceClip to destCollection, depending on the value of 'action' """
    contin = True
    if action == 'Copy':
       # Make a duplicate of the clip to be copied
       newClip = sourceClip.duplicate()
       # To place the copy in the destination collection, alter its Collection Number, Collection ID, and Sort Order value
       newClip.collection_num = destCollection.number
       newClip.collection_id = destCollection.id
    elif action == 'Move':
        try:
           # Lock the Clip Record to prevent other users from altering it simultaneously
           sourceClip.lock_record()
           # To move a clip, alter its Collection Number, Collection ID, and Sort Order value
           sourceClip.collection_num = destCollection.number
           sourceClip.collection_id = destCollection.id
        except TransanaExceptions.RecordLockedError, e:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('You cannot move Clip "%s"') + \
                                 _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
            else:
                prompt = _('You cannot move Clip "%s"') + \
                         _('.\nThe record is currently locked by %s.\nPlease try again later.')
            errordlg = Dialogs.ErrorDialog(None, prompt % (sourceClip.id, e.user))
            errordlg.ShowModal()
            errordlg.Destroy()
            contin = False
    if contin:
       # NOTE:  CopyMoveClip places the copy at the end of the Collection's Clip List.  If that's not
       #        what we want, we can call ChangeClipOrder later.

       # Get the highest SortOrder value and add one to it 
       clipCount = DBInterface.getMaxSortOrder(destCollection.number) + 1

       # Check for Duplicate Clip Names, an error condition
       # First, get the name of the appropriate Clip Object
       if action == 'Copy':
          clipName = newClip.id
       elif action == 'Move':
          clipName = sourceClip.id
       # See if the Clip Name already exists in the Destination Collection
       (dupResult, newClipName) = CheckForDuplicateObjName(clipName, 'ClipNode', treeCtrl, destNode)
       
       # If a Duplicate Clip Name is found and the error situation not resolved, show an Error Message
       if dupResult:
          # Unlock the source clip (before presenting the dialog to keep if from being locked by a slow user response.)
          if action == 'Move':
              sourceClip.unlock_record()
          # Report the failure to the user, although it's already known to have failed because they pressed "cancel".
          if 'unicode' in wx.PlatformInfo:
              # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
              prompt = unicode(_('%s cancelled for Clip "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceClip.id)
          else:
              prompt = _('%s cancelled for Clip "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceClip.id)
          dlg = Dialogs.ErrorDialog(treeCtrl, prompt)
          dlg.ShowModal()
          dlg.Destroy()
          return None
       else:
          # The user may have given the clip a new Clip Name in CheckForDuplicateObjName.  
          if newClipName != clipName:
             # If so, use this new name!
             if action == 'Copy':
                newClip.id = newClipName
             elif action == 'Move':
                 sourceClip.id = newClipName 

          if action == 'Copy':
             # Now that we know the number of clips in the collection, assign that as sortOrder
             newClip.sort_order = clipCount
             # Save the new Clip to the database.
             newClip.db_save()
          elif action == 'Move':
             # Now that we know the number of clips in the collection, assign that as sortOrder
             sourceClip.sort_order = clipCount
             # Save the new Clip to the database.
             sourceClip.db_save()
             # Unlock the Clip Record
             sourceClip.unlock_record()
                  
             # Remove the old Clip from the Tree.
             # delete_Node needs to be able to climb the tree, so we need to build the Node List that
             # tells it what to delete.  Start with the sourceCollection.
             nodeList = (sourceCollection.id,)
             # Load the Current Collection so we can find out about its parent, and work backwards from here.
             tempCollection = Collection.Collection(sourceCollection.number)
             # While the Current Collection has a defined parent...
             while tempCollection.parent > 0:
                # Make the Parent the Current Collection
                tempCollection = Collection.Collection(tempCollection.parent)
                # Add the Parent (now the Current Collection) to the FRONT of the Node List
                nodeList = (tempCollection.id,) + nodeList
             # Now add the Collections Root to the front of the Node List and the Clip's original name to the end of the Node List
             nodeList = (_('Collections'), ) + nodeList + (clipName, )
             # Now request that the defined node be deleted
             treeCtrl.delete_Node(nodeList, 'ClipNode')
             
             # Check for Keyword Examples that need to also be renamed!
             for (kwg, kw, clipNumber, clipID) in DBInterface.list_all_keyword_examples_for_a_clip(sourceClip.number):
                 nodeList = (_('Keywords'), kwg, kw, clipName)
                 exampleNode = treeCtrl.select_Node(nodeList, 'KeywordExampleNode')
                 treeCtrl.SetItemText(exampleNode, newClipName)
                 # If we're in the Multi-User mode, we need to send a message about the change
                 if not TransanaConstants.singleUserVersion:
                     # Begin constructing the message with the old and new names for the node
                     msg = " >|< %s >|< %s" % (clipName, newClipName)
                     # Get the full Node Branch by climbing it to two levels above the root
                     while (treeCtrl.GetItemParent(treeCtrl.GetItemParent(exampleNode)) != treeCtrl.GetRootItem()):
                         # Update the selected node indicator
                         exampleNode = treeCtrl.GetItemParent(exampleNode)
                         # Prepend the new Node's name on the Message with the appropriate seperator
                         msg = ' >|< ' + treeCtrl.GetItemText(exampleNode) + msg
                     # The first parameter is the Node Type.  The second one is the UNTRANSLATED root node.
                     # This must be untranslated to avoid problems in mixed-language environments.
                     # Prepend these on the Messsage
                     msg = "KeywordExampleNode >|< Keywords" + msg
                     # Send the Rename Node message
                     if TransanaGlobal.chatWindow != None:
                         TransanaGlobal.chatWindow.SendMessage("RN %s" % msg)
                
             # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
             ClearClipboard()

          # Add the new Clip to the Database Tree Tab
          # To add a Clip, we need to build the node list for the tree's add_Node method to climb.
          # We need to add all of the Collection Parents to our Node List, so we'll start by loading
          # the current Collection
          tempCollection = Collection.Collection(destCollection.number)
          # Add the current Collection Name, and work backwards from here.
          nodeList = (tempCollection.id,)
          # Repeat this process as long as the Collection we're looking at has a defined Parent...
          while tempCollection.parent > 0:
             # Load the Parent Collection
             tempCollection = Collection.Collection(tempCollection.parent)
             # Add this Collection's name to the FRONT of the Node List
             nodeList = (tempCollection.id,) + nodeList
          # Now add the Collections Root node to the front of the Node List and the
          # Clip Name to the back of the Node List
          if action == 'Copy':
             nodeList = (_('Collections'), ) + nodeList + (newClip.id, )
             # Add the Node to the Tree
             treeCtrl.add_Node('ClipNode', nodeList, newClip.number, newClip.collection_num, sortOrder=newClip.sort_order)

             # Now let's communicate with other Transana instances if we're in Multi-user mode
             if not TransanaConstants.singleUserVersion:
                msg = "ACl %s"
                data = (nodeList[1],)

                for nd in nodeList[2:]:
                   msg += " >|< %s"
                   data += (nd, )
                if TransanaGlobal.chatWindow != None:
                   TransanaGlobal.chatWindow.SendMessage(msg % data)

             # Now let's see if the clip had clip notes to copy!
             clipNoteList = DBInterface.list_of_notes(Clip=sourceClip.number, includeNumber=True)
             # For each note in the note list ...
             for clipNote in clipNoteList:
                 # Open a temporary copy of the note
                 tmpNote = Note.Note(clipNote[0])
                 # Duplicate the note (which automatically gives it an object number = 0)
                 newNote = tmpNote.duplicate()
                 # Assign the duplicate to the NEW clip
                 newNote.clip_num = newClip.number
                 # Save the new note
                 newNote.db_save()
                 # Add the new Note Node to the Tree
                 treeCtrl.add_Node('ClipNoteNode', nodeList + (newNote.id,), newNote.number, newClip.number)

                 # If the Notes Browser is open ...
                 if treeCtrl.parent.ControlObject.NotesBrowserWindow != None:
                     # ... Add the new note to the Notes Browser
                     treeCtrl.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('A', newNote)
                
                 # Now let's communicate with other Transana instances if we're in Multi-user mode
                 if not TransanaConstants.singleUserVersion:
                     # Prepare an Add Clip Note message
                     msg = "AClN Collections >|< %s"
                     # Convert the Node List to the form needed for messaging
                     data = (nodeList[1],)
                     for nd in (nodeList[2:] + (newNote.id,)):
                          msg += " >|< %s"
                          data += (nd, )
                     # Send the message
                     if TransanaGlobal.chatWindow != None:
                         TransanaGlobal.chatWindow.SendMessage(msg % data)

             # When copying a Clip and setting its sort order, we need to keep working with the new clip
             # rather than the old one.  Having CopyMoveClip return the new clip makes this easy.
             return newClip
          elif action == 'Move':
             nodeList = (_('Collections'), ) + nodeList + (sourceClip.id, )
             # Add the Node to the Tree
             treeCtrl.add_Node('ClipNode', nodeList, sourceClip.number, sourceClip.collection_num, sortOrder=sourceClip.sort_order)
             # If we are moving a Clip, the clip's Notes need to travel with the Clip.  The first step is to
             # get a list of those Notes.
             noteList = DBInterface.list_of_notes(Clip=sourceClip.number)
             # If there are Clip Notes, we need to make sure they travel with the Clip
             if noteList != []:
                 newNode = treeCtrl.select_Node(nodeList, 'ClipNode')
                 # We accomplish this using the TreeCtrl's "add_note_nodes" method
                 treeCtrl.add_note_nodes(noteList, newNode, Clip=sourceClip.number)
                 treeCtrl.Refresh()

             # Now let's communicate with other Transana instances if we're in Multi-user mode
             if not TransanaConstants.singleUserVersion:
                msg = "ACl %s"
                data = (nodeList[1],)

                for nd in nodeList[2:]:
                   msg += " >|< %s"
                   data += (nd, )
                if TransanaGlobal.chatWindow != None:
                   TransanaGlobal.chatWindow.SendMessage(msg % data)

             # When copying a Clip and setting its sort order, we need to keep working with the new clip
             # rather than the old one.  Having CopyMoveClip return the new clip makes this easy.
             return sourceClip

def CopyMoveSnapshot(treeCtrl, destNode, sourceSnapshot, sourceCollection, destCollection, action):
    """ This function copies or moves sourceSnapshot to destCollection, depending on the value of 'action' """
    contin = True
    if action == 'Copy':
       # Make a duplicate of the Snapshot to be copied
       newSnapshot = sourceSnapshot.duplicate()
       # To place the copy in the destination collection, alter its Collection Number, Collection ID, and Sort Order value
       newSnapshot.collection_num = destCollection.number
       newSnapshot.collection_id = destCollection.id
    elif action == 'Move':
        try:
           # Lock the Snapshot Record to prevent other users from altering it simultaneously
           sourceSnapshot.lock_record()
           # To move a Snapshot, alter its Collection Number, Collection ID, and Sort Order value
           sourceSnapshot.collection_num = destCollection.number
           sourceSnapshot.collection_id = destCollection.id
        except TransanaExceptions.RecordLockedError, e:
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('You cannot move Snapshot "%s"') + \
                                 _('.\nThe record is currently locked by %s.\nPlease try again later.'), 'utf8')
            else:
                prompt = _('You cannot move Snapshot "%s"') + \
                         _('.\nThe record is currently locked by %s.\nPlease try again later.')
            errordlg = Dialogs.ErrorDialog(None, prompt % (sourceSnapshot.id, e.user))
            errordlg.ShowModal()
            errordlg.Destroy()
            contin = False
    if contin:
       # NOTE:  CopyMoveSnapshot places the copy at the end of the Collection's Object List.  If that's not
       #        what we want, we can call ChangeClipOrder later.

       # Get the highest SortOrder value and add one to it 
       objCount = DBInterface.getMaxSortOrder(destCollection.number) + 1

       # Check for Duplicate Snapshot Names, an error condition
       # First, get the name of the appropriate Snapshot Object
       if action == 'Copy':
          snapshotName = newSnapshot.id
       elif action == 'Move':
          snapshotName = sourceSnapshot.id
       # See if the Snapshot Name already exists in the Destination Collection
       (dupResult, newSnapshotName) = CheckForDuplicateObjName(snapshotName, 'SnapshotNode', treeCtrl, destNode)

       # If a Duplicate Snapshot Name is found and the error situation not resolved, show an Error Message
       if dupResult:
          # Unlock the source snapshot (before presenting the dialog to keep if from being locked by a slow user response.)
          if action == 'Move':
              sourceSnapshot.unlock_record()
          # Report the failure to the user, although it's already known to have failed because they pressed "cancel".
          if 'unicode' in wx.PlatformInfo:
              # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
              prompt = unicode(_('%s cancelled for Snapshot "%s".  Duplicate Item Name Error.'), 'utf8') % (copyMovePrompt, sourceSnapshot.id)
          else:
              prompt = _('%s cancelled for Snapshot "%s".  Duplicate Item Name Error.') % (copyMovePrompt, sourceSnapshot.id)
          dlg = Dialogs.ErrorDialog(treeCtrl, prompt)
          dlg.ShowModal()
          dlg.Destroy()
          return None
       else:
          # The user may have given the snapshot a new Snapshot Name in CheckForDuplicateObjName.  
          if newSnapshotName != snapshotName:
             # If so, use this new name!
             if action == 'Copy':
                newSnapshot.id = newSnapshotName
             elif action == 'Move':
                 sourceSnapshot.id = newSnapshotName 

          if action == 'Copy':
             # Now that we know the number of objects in the collection, assign that as sortOrder
             newSnapshot.sort_order = objCount
             # Save the new Snapshot to the database.
             newSnapshot.db_save()
          elif action == 'Move':
             # Now that we know the number of objects in the collection, assign that as sortOrder
             sourceSnapshot.sort_order = objCount
             # Save the new Snapshot to the database.
             sourceSnapshot.db_save()
             # Unlock the Snapshot Record
             sourceSnapshot.unlock_record()

             # Remove the old Snapshot from the Tree.
             # delete_Node needs to be able to climb the tree, so we need to build the Node List that
             # tells it what to delete.  Start with the sourceCollection.
             nodeList = (sourceCollection.id,)
             # Load the Current Collection so we can find out about its parent, and work backwards from here.
             tempCollection = Collection.Collection(sourceCollection.number)
             # While the Current Collection has a defined parent...
             while tempCollection.parent > 0:
                # Make the Parent the Current Collection
                tempCollection = Collection.Collection(tempCollection.parent)
                # Add the Parent (now the Current Collection) to the FRONT of the Node List
                nodeList = (tempCollection.id,) + nodeList
             # Now add the Collections Root to the front of the Node List and the Snapshot's original name to the end of the Node List
             nodeList = (_('Collections'), ) + nodeList + (snapshotName, )
             # Now request that the defined node be deleted
             treeCtrl.delete_Node(nodeList, 'SnapshotNode')
             
             # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
             ClearClipboard()

          # Add the new Snapshot to the Database Tree Tab
          # To add a Snapshot, we need to build the node list for the tree's add_Node method to climb.
          # We need to add all of the Collection Parents to our Node List, so we'll start by loading
          # the current DESTINATION Collection
          tempCollection = Collection.Collection(destCollection.number)
          # Add the current Collection Name, and work backwards from here.
          nodeList = (tempCollection.id,)
          # Repeat this process as long as the Collection we're looking at has a defined Parent...
          while tempCollection.parent > 0:
             # Load the Parent Collection
             tempCollection = Collection.Collection(tempCollection.parent)
             # Add this Collection's name to the FRONT of the Node List
             nodeList = (tempCollection.id,) + nodeList
          # Now add the Collections Root node to the front of the Node List and the
          # Snapshot Name to the back of the Node List
          if action == 'Copy':
             nodeList = (_('Collections'), ) + nodeList + (newSnapshot.id, )
             # Add the Node to the Tree
             treeCtrl.add_Node('SnapshotNode', nodeList, newSnapshot.number, newSnapshot.collection_num, sortOrder=newSnapshot.sort_order)

             # Now let's communicate with other Transana instances if we're in Multi-user mode
             if not TransanaConstants.singleUserVersion:
                msg = "ASnap %s"
                data = (nodeList[1],)

                for nd in nodeList[2:]:
                   msg += " >|< %s"
                   data += (nd, )
                if TransanaGlobal.chatWindow != None:
                   TransanaGlobal.chatWindow.SendMessage(msg % data)

             # Now let's see if the snapshot had snapshot notes to copy!
             snapshotNoteList = DBInterface.list_of_notes(Snapshot=sourceSnapshot.number, includeNumber=True)
             # For each note in the note list ...
             for snapshotNote in snapshotNoteList:
                 # Open a temporary copy of the note
                 tmpNote = Note.Note(snapshotNote[0])
                 # Duplicate the note (which automatically gives it an object number = 0)
                 newNote = tmpNote.duplicate()
                 # Assign the duplicate to the NEW snapshot
                 newNote.snapshot_num = newSnapshot.number
                 # Save the new note
                 newNote.db_save()
                 # Add the new Note Node to the Tree
                 treeCtrl.add_Node('SnapshotNoteNode', nodeList + (newNote.id,), newNote.number, newSnapshot.number)

                 # If the Notes Browser is open ...
                 if treeCtrl.parent.ControlObject.NotesBrowserWindow != None:
                     # ... Add the new note to the Notes Browser
                     treeCtrl.parent.ControlObject.NotesBrowserWindow.UpdateTreeCtrl('A', newNote)

                 # Now let's communicate with other Transana instances if we're in Multi-user mode
                 if not TransanaConstants.singleUserVersion:
                     # Prepare an Add Snapshot Note message
                     msg = "ASnN Collections >|< %s"
                     # Convert the Node List to the form needed for messaging
                     data = (nodeList[1],)
                     for nd in (nodeList[2:] + (newNote.id,)):
                          msg += " >|< %s"
                          data += (nd, )
                     # Send the message
                     if TransanaGlobal.chatWindow != None:
                         TransanaGlobal.chatWindow.SendMessage(msg % data)

             # When copying a Snapshot and setting its sort order, we need to keep working with the new snapshot
             # rather than the old one.  Having CopyMoveSnapshot return the new snapshot makes this easy.
             return newSnapshot
          elif action == 'Move':
             nodeList = (_('Collections'), ) + nodeList + (sourceSnapshot.id, )
             # Add the Node to the Tree
             treeCtrl.add_Node('SnapshotNode', nodeList, sourceSnapshot.number, sourceSnapshot.collection_num, sortOrder=sourceSnapshot.sort_order)
             # If we are moving a Snapshot, the Snapshot's Notes need to travel with the Snapshot.  The first step is to
             # get a list of those Notes.
             noteList = DBInterface.list_of_notes(Snapshot=sourceSnapshot.number)
             # If there are Snapshot Notes, we need to make sure they travel with the Snapshot
             if noteList != []:
                 newNode = treeCtrl.select_Node(nodeList, 'SnapshotNode')
                 # We accomplish this using the TreeCtrl's "add_note_nodes" method
                 treeCtrl.add_note_nodes(noteList, newNode, Snapshot=sourceSnapshot.number)
                 treeCtrl.Refresh()

             # Now let's communicate with other Transana instances if we're in Multi-user mode
             if not TransanaConstants.singleUserVersion:
                msg = "ASnap %s"
                data = (nodeList[1],)

                for nd in nodeList[2:]:
                   msg += " >|< %s"
                   data += (nd, )
                if TransanaGlobal.chatWindow != None:
                   TransanaGlobal.chatWindow.SendMessage(msg % data)

             # When copying a Snapshot and setting its sort order, we need to keep working with the new snapshot
             # rather than the old one.  Having CopyMoveShapshot return the new snapshot makes this easy.
             return sourceSnapshot

def ChangeClipOrder(treeCtrl, destNode, sourceObject, sourceCollection):
    """ This function changes the order of the Items in a Collection """
    # Get the Destination Node Data
    destData = treeCtrl.GetPyData(destNode)
    # Get the Sort Order value for where we want the new item
    targetSortOrder = destData.sortOrder
    # If we can't lock all the clips in the collection, sort orders get all screwed up.
    # ... Set up a variable that signals failure
    allObjectsLocked = True

    # Create a Dictionary to hold all the Quote data, so we only need to have one copy of the quote
    Quotes = {}
    # Get all the quotes for the Source Collection Number
    quoteLockList = DBInterface.list_of_quotes_by_collectionnum(sourceCollection.number)
    # Start Exception Handling
    try:
        # For each Quote in the Collection ...
        for (tmpQuoteNum, tmpQuoteID, tmpCollectNum) in quoteLockList:
            # Load the Quote.
            tmpQuote = Quote.Quote(num=tmpQuoteNum)
            # Lock the QuoteRecord
            tmpQuote.lock_record()
            # If we have the Source Object, clear the sort_order value
            if isinstance(sourceObject, Quote.Quote) and (sourceObject.number == tmpQuoteNum):
                tmpQuote.sort_order = 0
            # Add this clip to the Quotes dictionary
            Quotes[tmpQuoteNum] = tmpQuote
    # If we couldn't get a lock on one or more of the quotes...
    except TransanaExceptions.RecordLockedError, e:
        # Set the "Failure" flag
        allObjectsLocked = False
        # Create an error message for the user
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            msg = unicode(_('Items in Collection "%s" are not in the desired order.') + '\n\n' + \
                          _('Transana could not change the sort order because you cannot obtain a lock on Quote "%s"') + \
                          _('.\nThe record is currently locked by %s.'), 'utf8')
        else:
            msg = _('Items in Collection "%s" are not in the desired order.') + '\n\n' + \
                  _('Transana could not change the sort order because you cannot obtain a lock on Quote "%s"') + \
                  _('.\nThe record is currently locked by %s.')
        # Display the error message
        dlg = Dialogs.ErrorDialog(None, msg % (sourceCollection.id, tmpQuote.id, e.user))
        dlg.ShowModal()
        dlg.Destroy()

        # If the Change of Sort Orders failed, put the new object at the END
        targetSortOrder = DBInterface.getMaxSortOrder(destData.parent) + 1

    # Create a Dictionary to hold all the Clip data, so we only need to have one copy of the clip
    Clips = {}
    # Get all the clips for the Source Collection Number
    clipLockList = DBInterface.list_of_clips_by_collectionnum(sourceCollection.number)
    # Start Exception Handling
    try:
        # For each Clip in the Collection ...
        for (tmpClipNum, tmpClipID, tmpCollectNum) in clipLockList:
            # Load the Clip.
            tmpClip = Clip.Clip(tmpClipNum)
            # Lock the Clip Record
            tmpClip.lock_record()
            # If we have the Source Object, clear the sort_order value
            if isinstance(sourceObject, Clip.Clip) and (sourceObject.number == tmpClipNum):
                tmpClip.sort_order = 0
            # Add this clip to the Clips dictionary
            Clips[tmpClipNum] = tmpClip
    # If we couldn't get a lock on one or more of the clips ...
    except TransanaExceptions.RecordLockedError, e:
        # Set the "Failure" flag
        allObjectsLocked = False
        # Create an error message for the user
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            msg = unicode(_('Items in Collection "%s" are not in the desired order.') + '\n\n' + \
                          _('Transana could not change the sort order because you cannot obtain a lock on Clip "%s"') + \
                          _('.\nThe record is currently locked by %s.'), 'utf8')
        else:
            msg = _('Items in Collection "%s" are not in the desired order.') + '\n\n' + \
                  _('Transana could not change the sort order because you cannot obtain a lock on Clip "%s"') + \
                  _('.\nThe record is currently locked by %s.')
        # Display the error message
        dlg = Dialogs.ErrorDialog(None, msg % (sourceCollection.id, tmpClip.id, e.user))
        dlg.ShowModal()
        dlg.Destroy()

        # If the Change of Sort Orders failed, put the new object at the END
        targetSortOrder = DBInterface.getMaxSortOrder(destData.parent) + 1

    # Create a Dictionary to hold all the Snapshot data, so we only need to have one copy of the Snapshot
    Snapshots = {}
    # Get all the snapshots for the Source Collection Number
    snapshotLockList = DBInterface.list_of_snapshots_by_collectionnum(sourceCollection.number)
    # Start Exception Handling
    try:
        # For each Snapshot in the Collection ...
        for (tmpSnapshotNum, tmpSnapshotID, tmpCollectNum) in snapshotLockList:
            # Load the Snapshot.
            tmpSnapshot = Snapshot.Snapshot(tmpSnapshotNum)
            # Lock the Snapshot Record
            tmpSnapshot.lock_record()
            # If we have the Source Object, clear the sort_order value
            if isinstance(sourceObject, Snapshot.Snapshot) and (sourceObject.number == tmpSnapshotNum):
                tmpSnapshot.sort_order = 0
            # Add this snapshot to the Snapshots dictionary
            Snapshots[tmpSnapshotNum] = tmpSnapshot
    # If we couldn't get a lock on one or more of the clips ...
    except TransanaExceptions.RecordLockedError, e:
        # Set the "Failure" flag
        allObjectsLocked = False
        # Create an error message for the user
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            msg = unicode(_('Items in Collection "%s" are not in the desired order.') + '\n\n' + \
                          _('Transana could not change the sort order because you cannot obtain a lock on Snapshot "%s"') + \
                          _('.\nThe record is currently locked by %s.'), 'utf8')
        else:
            msg = _('Items in Collection "%s" are not in the desired order.') + '\n\n' + \
                  _('Transana could not change the sort order because you cannot obtain a lock on Snapshot "%s"') + \
                  _('.\nThe record is currently locked by %s.')
        # Display the error message
        dlg = Dialogs.ErrorDialog(None, msg % (sourceCollection.id, tmpSnapshot.id, e.user))
        dlg.ShowModal()
        dlg.Destroy()

        # If the Change of Sort Orders failed, put the new object at the END
        targetSortOrder = DBInterface.getMaxSortOrder(destData.parent) + 1

    # If locking ALL clips DID NOT fail ...
    if allObjectsLocked:

        # Iterate through the Clips and Snapshots.  Increase SortOrder for everything above the desired position.
        # Sort the Collection.  Then update the PyData for each entry!!
        # This should be significantly faster!

        if DEBUG:
            print
            print "DragAndDropObjects.ChangeClipOrder():", sourceObject.number, sourceObject.sort_order

        # Iterate through the Quotes dictionary
        for tmpQuoteNum in Quotes.keys():

            if DEBUG:
                print "Quotes", tmpQuoteNum, Quotes[tmpQuoteNum].sort_order, '-->',
            
            # If the quote has a sort_order of 0, it hasn't been assigned
            if Quotes[tmpQuoteNum].sort_order == 0:
                # ... which means it's the one just inserted
                Quotes[tmpQuoteNum].sort_order = targetSortOrder
                # Save the Quote
                Quotes[tmpQuoteNum].db_save()
            # if the Quote's Sort Order value is greater than the Target Sort Order ...
            elif Quotes[tmpQuoteNum].sort_order >= targetSortOrder:
                # Change the Sort Order value for each Quote
                Quotes[tmpQuoteNum].sort_order += 1
                # Save the Quote
                Quotes[tmpQuoteNum].db_save()

            if DEBUG:
                print Quotes[tmpQuoteNum].sort_order
            
        # Iterate through the Clips dictionary
        for tmpClipNum in Clips.keys():

            if DEBUG:
                print "Clips", tmpClipNum, Clips[tmpClipNum].sort_order, '-->',
            
            # If the clip has a sort_order of 0, it hasn't been assigned
            if Clips[tmpClipNum].sort_order == 0:
                # ... which means it's the one just inserted
                Clips[tmpClipNum].sort_order = targetSortOrder
                # Save the Clip
                Clips[tmpClipNum].db_save()
            # if the Clip's Sort Order value is greater than the Target Sort Order ...
            elif Clips[tmpClipNum].sort_order >= targetSortOrder:
                # Change the Sort Order value for each Clip
                Clips[tmpClipNum].sort_order += 1
                # Save the Clip
                Clips[tmpClipNum].db_save()

            if DEBUG:
                print Clips[tmpClipNum].sort_order
            
        # Iterate through the Snapshots dictionary
        for tmpSnapshotNum in Snapshots.keys():

            if DEBUG:
                print "Snapshots", tmpSnapshotNum, Snapshots[tmpSnapshotNum].sort_order, '-->',
            
            # If the Snapshot has a sort_order of 0, it hasn't been assigned
            if Snapshots[tmpSnapshotNum].sort_order == 0:
                # ... which means it's the one just inserted
                Snapshots[tmpSnapshotNum].sort_order = targetSortOrder
                # Save the Clip
                Snapshots[tmpSnapshotNum].db_save()
            # if the Snapshot's Sort Order value is greater than the Target Sort Order ...
            elif Snapshots[tmpSnapshotNum].sort_order >= targetSortOrder:
                # Change the Sort Order value for each Snapshot
                Snapshots[tmpSnapshotNum].sort_order += 1
                # Save the Snapshot
                Snapshots[tmpSnapshotNum].db_save()

            if DEBUG:
                print Snapshots[tmpSnapshotNum].sort_order
            
        # Assign the Sort Order for the object passed in
        sourceObject.sort_order = targetSortOrder

        if DEBUG:
            print "sourceObject:", sourceObject.number, sourceObject.sort_order
            print

        # Save the Source Object
        sourceObject.db_save()

        # Now Iterate through all the Collection's children, updating the SortOrder in the PyData
        
        # First, let's identify the Parent Node we're working with
        parentNode = treeCtrl.GetItemParent(destNode)
        # wxTreeCtrl requires the "cookie" value to list children.  Initialize it.
        cookie = 0
        # Get the first child of the dropNode 
        (tempNode, cookie) = treeCtrl.GetFirstChild(parentNode)

        if DEBUG:
            print "updating NODES:"
        
        # Iterate through all the dropNode's children
        while tempNode.IsOk():
            # Get the current child's Node Data
            tempNodeData = treeCtrl.GetPyData(tempNode)
            if tempNodeData.nodetype == 'QuoteNode':
                # update the node's Sort Order
                tempNodeData.sortOrder = Quotes[tempNodeData.recNum].sort_order
            elif tempNodeData.nodetype == 'ClipNode':
                # update the node's Sort Order
                tempNodeData.sortOrder = Clips[tempNodeData.recNum].sort_order
            elif tempNodeData.nodetype == 'SnapshotNode':
                # update the node's Sort Order
                tempNodeData.sortOrder = Snapshots[tempNodeData.recNum].sort_order
            # Save the Node Data to the TreeCtrl Node
            treeCtrl.SetPyData(tempNode, tempNodeData)

            if DEBUG:
                print treeCtrl.GetItemText(tempNode), tempNodeData.sortOrder

            # If we are looking at the last Child in the Parent's Node, exit the while loop
            if tempNode == treeCtrl.GetLastChild(parentNode):
                break
            # If not, load the next Child record
            else:
                (tempNode, cookie) = treeCtrl.GetNextChild(parentNode, cookie)

        if DEBUG:
            print
            print "Sorting children"
            print
            print
        
        # Sort the Database Tree's Collection Node
        treeCtrl.SortChildren(treeCtrl.GetItemParent(destNode))

    # If Sort Order failed ...
    else:
        # Assign the Sort Order for the object passed in
        sourceObject.sort_order = targetSortOrder
        # Save the Source Object
        sourceObject.db_save()
        # ... reset the order of objects in the node, no need to send MU messages
        treeCtrl.UpdateCollectionSortOrder(treeCtrl.GetItemParent(destNode), sendMessage=False)

        if isinstance(sourceObject, Snapshot.Snapshot):
            nodeList = sourceObject.GetNodeData(False)

#            print "DragAndDropObjects.ChangeClipOrder():", nodeList
            
            # Now let's communicate with other Transana instances if we're in Multi-user mode
            if not TransanaConstants.singleUserVersion:
                msg = "OC %s"
                data = (_('Collections'), )

                for nd in nodeList[0:]:
                    msg += " >|< %s"
                    data += (nd, )

#                print '"msg %s"' % (data,)
#                print
                
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage(msg % data)

    # Iterate through the Quotes dictionary
    for tmpQuoteNum in Quotes.keys():
        # Unlock each of the locked Quotes
        Quotes[tmpQuoteNum].unlock_record()
    # Iterate through the Clips dictionary
    for tmpClipNum in Clips.keys():
        # Unlock each of the locked Clips
        Clips[tmpClipNum].unlock_record()
    # Iterate through the Snapshots dictionary
    for tmpSnapshotNum in Snapshots.keys():
        # Unlock each of the locked Snapshots
        Snapshots[tmpSnapshotNum].unlock_record()
    # Return the flag that indicates success or failure
    return allObjectsLocked

def CreateQuickClip(clipData, kwg, kw, dbTree, extraKeywords=[]):
    """ Create a "Quick Clip", which is the implementation of a simplified form of Clip Creation """
    # We need to error check to make sure we have a legal Clip spec.
    # If multi-transcript QuickClip is being created, clipData.text will be a CLIP, not Transcript text!!!!
    if (clipData.clipStart >= clipData.clipStop) or \
       (not isinstance(clipData.text, Clip.Clip) and (clipData.text == "")):
        msg = _("You must select some text in the Transcript to be able to create a Quick Clip.")
        errorDlg = Dialogs.ErrorDialog(None, msg)
        errorDlg.ShowModal()
        errorDlg.Destroy()
    else:
        # Load the Episode that is the source of the current selection.
        sourceEpisode = Episode.Episode(clipData.episodeNum)

        # Check to see if the clip starts before the media file starts (due to Adjust Indexes)
        if clipData.clipStart < 0.0:
            prompt = _('The starting point for a Clip cannot be before the start of the media file.')
            errordlg = Dialogs.ErrorDialog(None, prompt)
            errordlg.ShowModal()
            errordlg.Destroy()
            # If so, cancel the clip creation
            return

        # Check to see if the clip starts after the media file ends (due to Adjust Indexes)
        if clipData.clipStart >= sourceEpisode.tape_length:
            prompt = _('The starting point for a Clip cannot be after the end of the media file.')
            errordlg = Dialogs.ErrorDialog(None, prompt)
            errordlg.ShowModal()
            errordlg.Destroy()
            # If so, cancel the Clip creation
            return

        # Check to see if the clip goes all the way to the end of the media file.  If so, it's probably an accident
        # due to the lack of an ending time code.  We'll skip this in the last 30 seconds of the media file, though.
        if (clipData.clipStop == sourceEpisode.tape_length) and \
           (clipData.clipStop - clipData.clipStart > 30000):
            prompt = _('The ending point for this Clip is the end of the media file.  Do you want to create this clip?')
            errordlg = Dialogs.QuestionDialog(None, prompt, _("Transana Error"))
            result = errordlg.LocalShowModal()
            errordlg.Destroy()
            # If the user says NO, they don't want to create it ...
            if result == wx.ID_NO:
                # .. cancel Clip Creation
                return

        # Check to see if the clip ends after the media file ends (due to Adjust Indexes)
        if clipData.clipStop > sourceEpisode.tape_length:
            prompt = _('The ending point for this Clip is after the end of the media file.  This clip may not end where you expect.')
            errordlg = Dialogs.ErrorDialog(None, prompt)
            errordlg.ShowModal()
            errordlg.Destroy()
            # We don't cancel clip creation, but we do adjust the end of the clip.
            clipData.clipStop = sourceEpisode.tape_length

        # Let's check to see if there's an appropriate Collection for the Quick Clips
        (collectNum, collectName, newCollection) = DBInterface.locate_quick_quotes_and_clips_collection()
        # If a new collection was created, ...
        if newCollection:
            # ... we need to add it to the database tree.
            # Build the node data needed to add the collection.
            nodeData = (_('Collections'), collectName)
            # Add the new Collection to the data tree
            dbTree.add_Node('CollectionNode', nodeData, collectNum, 0, expandNode=False)
            # If in multi-user mode ...
            if not TransanaConstants.singleUserVersion:
                # ... inform the Message Server that a Collection has been added
                msg = "AC %s"
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage(msg % collectName)

        # If we have a single-transcript QuickClip, then clipData.transcriptNum is > 0 and clipData.text is RTF text.
        # If we have a multi-transcript QuickClip, then clipData.transcriptNum == 0 and clipData.text is actually a
        # Clip object.  In this case, we need to pass ALL source transcript numbers to the CheckForDuplicateQuickClip()
        # function.
        if clipData.transcriptNum == 0:
            # Initialize a list.
            clipData.transcriptNum = []
            # for each transcript ...
            for tr in clipData.text.transcripts:
                # ... append its source transcript to the list.
                clipData.transcriptNum.append(tr.source_transcript)
        # If we have a single media file SOURCE ...
        if clipData.videoCheckboxData == []:
            # ... then drop the media file in the list of video files
            vidFiles = [sourceEpisode.media_filename]
        # If we have a MULTIPLE media file SOURCE ...
        else:
            # ... initialize a list of media files ...
            vidFiles = []
            # ... if the FIRST media file is to be included ...
            if clipData.videoCheckboxData[0][0]:
                # ... add that to the media file list
                vidFiles.append(sourceEpisode.media_filename)
            # Initialize a counter
            cnt = 0
            # For each additional media file in the NEW CLIP media file checkboxes ...  (skip the original file, #0)
            for cbData in clipData.videoCheckboxData[1:]:
                # .. if the media file is checked to be included ...
                if cbData[0]:
                    # ... then get the file name from the EPISODE and add it to the media file list
                    vidFiles.append(sourceEpisode.additional_media_files[cnt]['filename'])
                # increment the counter
                cnt += 1
        # Check to see if a Quick Clip for this selection in this Transcript in this Episode has already been created.
        dupClipNum = DBInterface.CheckForDuplicateQuickClip(collectNum, clipData.episodeNum, clipData.transcriptNum, clipData.clipStart, clipData.clipStop, vidFiles)
        # -1 indicates no duplicate Quick Clip.  If there IS a duplicate ...
        if dupClipNum > -1:
            # ... load the existing Quick Clip
            quickClip = Clip.Clip(dupClipNum)
            # If we're supposed to warn users ...
            if TransanaGlobal.configData.quickClipWarning:
                # Inform the user of the duplication.
                msg = _('A Quick Clip matching this selection already exists.\nKeyword "%s : %s" will be added to\nQuick Clip "%s".')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8')
                tempDlg = Dialogs.InfoDialog(None, msg % (kwg, kw, quickClip.id))
                tempDlg.ShowModal()
                tempDlg.Destroy()
            # Attempt to add the selected keyword to the existing Quick Clip
            try:
                # Attempt to get a record lock
                quickClip.lock_record()
                # Add the keyword to the Clip record
                quickClip.add_keyword(kwg, kw)
                # If there are extra keywords ...
                for (kwg2, kw2) in extraKeywords:
                    # ... add them to the quick clip
                    quickClip.add_keyword(kwg2, kw2)
                # Save the Clip Record
                quickClip.db_save()
                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    msg = 'Clip %d' % quickClip.number
                    if TransanaGlobal.chatWindow != None:
                        # Send the "Update Keyword List" message
                        TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                # See if the Keyword visualization needs to be updated.
                dbTree.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Episode Keyword Visualization
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', quickClip.number, quickClip.episode_num))

                # Unlock the record
                quickClip.unlock_record()
            # Handle "RecordLockedError" exception
            except TransanaExceptions.RecordLockedError, e:
                TransanaExceptions.ReportRecordLockedException(_("Clip"), quickClip.id, e)
            # Handle "SaveError" exception
            except TransanaExceptions.SaveError:
                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
            # Handle other exceptions
            except:
                errordlg = Dialogs.ErrorDialog(None, "%s %s" % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
        # If there is NO duplicate Quick Clip record ...
        else:
            # Determine the next Quick Clip Name
            # Establish the base Clip Name.  Limit the size of the Episode Name so we don't overflow
            # the Clip ID field length in any language.
            # Start with the i18n version of "Quick Clip"
            if 'unicode' in wx.PlatformInfo:
                baseName = unicode(_('Quick Clip'), 'utf8') + ' '
            else:
                baseName = _('Quick Clip') + ' '
            # If multi-user ...
            if not TransanaConstants.singleUserVersion:
                # ... then pre-pend the username
                baseName = '- ' + TransanaGlobal.userName + ' - ' + baseName
            # pre-pent as much of the Episode name as will fit.
            baseName = sourceEpisode.id[:94 - len(baseName)] + ' ' + baseName
            # Start numbering at 1
            baseNum = 1
            # Get a list of all QuickClips
            clipList = DBInterface.list_of_clips_by_collectionnum(collectNum)
            # Iterate through the list
            for clipItem in clipList:
                # Isolate the Clip Name
                clipName = clipItem[1]
                # We can ignore errors in converting clip names to integers, but we have to trap it.
                try:
                    # make sure the Base Name matches the comparison Quick Clip name.
                    if baseName == clipName[:len(baseName)]:
                        # Get the integer portion of the Clip Name that follows the base Clip Name
                        clipNum = int(clipName[len(baseName):])
                        # We're looking for the largest value.
                        if clipNum >= baseNum:
                            # The base number value should always be 1 larger than the largest used value.
                            baseNum = clipNum + 1
                except:
                    pass

            # Create a Clip Object and populate it with the proper data
            quickClip = Clip.Clip()
            quickClip.id = baseName + str(baseNum)
            quickClip.collection_num = collectNum
            quickClip.episode_num = clipData.episodeNum

            # Handle multiple video data here.
            
            # Start the clip off with the Episode's offset, though this could change if the first video wasn't used!
            quickClip.offset = sourceEpisode.offset
            # Initially, assume that we don't need to shift the Clip offset, i.e. that the offset shift is ZERO
            offsetShift = 0
            # If there's no Video Checkbox data (ie no video checkboxes) or the first entry's "Include in Clip" option is selected ...
            if (clipData.videoCheckboxData == []) or (clipData.videoCheckboxData[0][0]):
                # The Clip's Media Filename comes from the Episode Record
                quickClip.media_filename = sourceEpisode.media_filename
            # Audio defaults to 1, but if there are multiple videos and the first video should NOT include audio ...
            if (clipData.videoCheckboxData != []) and (not clipData.videoCheckboxData[0][1]):
                # ... then we should set it to 0 (off)
                quickClip.audio = 0

            # For each video after the first one (which has already been handled) ...
            for x in range(1, len(clipData.videoCheckboxData)):
                # ... get the checkbox data
                (videoCheck, audioCheck) = clipData.videoCheckboxData[x]
                # If the video is supposed to be included in the Clip ...
                if videoCheck:
                    # if this is the FIRST video to be included, put the data in the Clip object.
                    if quickClip.media_filename == '':
                        # Grab the file name
                        quickClip.media_filename = sourceEpisode.additional_media_files[x - 1]['filename']
                        # If we wind up here, we need to shift the offset values.  Remember the amount to shift them.
                        offsetShift = sourceEpisode.additional_media_files[x - 1]['offset']
                        # Add the offset shift value to the Clip's gobal offset
                        quickClip.offset += offsetShift
                        # Note whether the audio should be played by default
                        quickClip.audio = audioCheck
                    # if this is NOT hte first video to be included, put the data in the additional_media_files structure,
                    # adjusting the offset by the offsetShift value if needed.  (YES, minus here, plus above.)
                    else:
                        quickClip.additional_media_files = {'filename' : sourceEpisode.additional_media_files[x - 1]['filename'],
                                                            'length'   : quickClip.clip_stop - quickClip.clip_start,
                                                            'offset'   : sourceEpisode.additional_media_files[x - 1]['offset'] - offsetShift,
                                                            'audio'    : audioCheck }
            # if NO media files were selected to be included, create an error message.
            if quickClip.media_filename == '':
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    prompt = unicode(_('Quick Clip Creation cancelled.  No media files have been selected for inclusion.'), 'utf8')
                else:
                    prompt = _('Quick Clip Creation cancelled.  No media files have been selected for inclusion.')
                errordlg = Dialogs.ErrorDialog(None, prompt)
                errordlg.ShowModal()
                errordlg.Destroy()
                # If Clip Creation fails, we don't need to continue any more.
                contin = False
                # Let's get out of here!
                return
    
            quickClip.clip_start = clipData.clipStart
            quickClip.clip_stop = clipData.clipStop
            quickClip.sort_order = DBInterface.getMaxSortOrder(collectNum) + 1

            # If we're creating a multi-transcript Quick Clip, the clipData.text field will actually be a Clip object!
            if isinstance(clipData.text, Clip.Clip):
                # If this is the case, just point the quick clip to the transcripts that are passed in.
                quickClip.transcripts = clipData.text.transcripts
            # If we're dealing with a single-transcript Quick Clip, we'll just have text here.
            else:
                # Create a Transcript object
                tempTranscript = Transcript.Transcript()
                # Get the Episode Number
                tempTranscript.episode_num = quickClip.episode_num
                # Get the Source Transcript number
                tempTranscript.source_transcript = clipData.transcriptNum
                # Grab the Clip Start and Stop times from the Clip object
                tempTranscript.clip_start = quickClip.clip_start
                tempTranscript.clip_stop = quickClip.clip_stop
                # Assign the Transcript Text
                if clipData.text == u'<(transcript-less clip)>':
                    tempTranscript.text = ''
                else:
                    tempTranscript.text = clipData.text
                # Add the Temporary Transcript to the Quick Clip
                quickClip.transcripts.append(tempTranscript)

            # Add the Episode Keywords as default Clip Keywords
            quickClip.keyword_list = sourceEpisode.keyword_list
            # Add the keyword that initiated the Quick Clip
            quickClip.add_keyword(kwg, kw)
            # If there are extra keywords ...
            for (kwg2, kw2) in extraKeywords:
                # ... add them to the quick clip
                quickClip.add_keyword(kwg2, kw2)

            # If we're in multi-user mode ...
            if not TransanaConstants.singleUserVersion:
                if 'unicode' in wx.PlatformInfo:
                    data = (unicode(_("Transana Users"), 'utf8'), TransanaGlobal.userName)
                else:
                    data = (_("Transana Users"), TransanaGlobal.userName)
                # ... determine if the Username is already a Keyword
                if DBInterface.check_username_as_keyword():
                    # Add the new Keyword to the data tree
                    dbTree.add_Node('KeywordNode', (_('Keywords'), _("Transana Users"), TransanaGlobal.userName), 0, _("Transana Users"), expandNode=False)
                    # Inform the Message Server of the added Keyword
                    if TransanaGlobal.chatWindow != None:
                        msg = "AK %s >|< %s"
                        if 'unicode' in wx.PlatformInfo:
                            msg = unicode(msg, 'utf8')
                        TransanaGlobal.chatWindow.SendMessage(msg % data)
                # Add the keyword to the Quick Clip
                quickClip.add_keyword(data[0], data[1])

            # If we're in Demo mode, an exception can be raised here if the Clip Limit is exceeded.
            # We need to trap that.
            try:
                # Save the Quick Clip
                quickClip.db_save()

                # We need to add the Clip to the database tree.
                # Build the node data needed to add the clip.
                nodeData = (_('Collections'), collectName, quickClip.id)
                # Add the new Collection to the data tree
                dbTree.add_Node('ClipNode', nodeData, quickClip.number, quickClip.collection_num, sortOrder=quickClip.sort_order, expandNode=False)
                # Inform the Message Server of the added Clip
                if not TransanaConstants.singleUserVersion:
                    msg = "ACl %s >|< %s"
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage(msg % (collectName, quickClip.id))

                # See if the Keyword visualization needs to be updated.
                dbTree.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Episode Keyword Visualization
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Clip', quickClip.number, quickClip.episode_num))
            # Process the SaveError.  This should ONLY occur in Demo Mode if the Clip Limit is exceeded.
            except TransanaExceptions.SaveError:
                # Remove the part of the error message that refers to clicking "Cancel" in the "Add Clip" dialog box!
                prompt = sys.exc_info()[1].reason.split('\n')[0]

                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, prompt)
                errordlg.ShowModal()
                errordlg.Destroy()

def CreateQuickQuote(quoteData, kwg, kw, dbTree, extraKeywords=[]):
    """ Create a "Quick Quote", which is the implementation of a simplified form of Quote Creation """
    # We need to error check to make sure we have a legal Quote spec.
    if (quoteData.startChar >= quoteData.endChar) or \
       (quoteData.text == ""):
        msg = _("You must select some text in the Document to be able to create a Quick Quote.")
        errorDlg = Dialogs.ErrorDialog(None, msg)
        errorDlg.ShowModal()
        errorDlg.Destroy()
    else:
        # Load the Document that is the source of the current selection.
        sourceDocument = Document.Document(quoteData.documentNum)

        # Let's check to see if there's an appropriate Collection for the Quick Clips
        (collectNum, collectName, newCollection) = DBInterface.locate_quick_quotes_and_clips_collection()
        # If a new collection was created, ...
        if newCollection:
            # ... we need to add it to the database tree.
            # Build the node data needed to add the collection.
            nodeData = (_('Collections'), collectName)
            # Add the new Collection to the data tree
            dbTree.add_Node('CollectionNode', nodeData, collectNum, 0, expandNode=False)
            # If in multi-user mode ...
            if not TransanaConstants.singleUserVersion:
                # ... inform the Message Server that a Collection has been added
                msg = "AC %s"
                if TransanaGlobal.chatWindow != None:
                    TransanaGlobal.chatWindow.SendMessage(msg % collectName)

        # Check to see if a Quick Quote for this selection in this Document has already been created.
        dupQuoteNum = DBInterface.CheckForDuplicateQuickQuote(collectNum, quoteData.documentNum, quoteData.startChar, quoteData.endChar)
        # -1 indicates no duplicate Quick Quote.  If there IS a duplicate ...
        if dupQuoteNum > -1:
            # ... load the existing Quick Quote
            quickQuote = Quote.Quote(dupQuoteNum)
            # If we're supposed to warn users ...
            if TransanaGlobal.configData.quickClipWarning:
                # Inform the user of the duplication.
                msg = _('A Quick Quote matching this selection already exists.\nKeyword "%s : %s" will be added to\nQuick Quote "%s".')
                if 'unicode' in wx.PlatformInfo:
                    # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                    msg = unicode(msg, 'utf8')
                tempDlg = Dialogs.InfoDialog(None, msg % (kwg, kw, quickQuote.id))
                tempDlg.ShowModal()
                tempDlg.Destroy()
            # Attempt to add the selected keyword to the existing Quick Quote
            try:
                # Attempt to get a record lock
                quickQuote.lock_record()
                # Add the keyword to the Quote record
                quickQuote.add_keyword(kwg, kw)
                # If there are extra keywords ...
                for (kwg2, kw2) in extraKeywords:
                    # ... add them to the Quick Quote
                    quickQuote.add_keyword(kwg2, kw2)
                # Save the Quote Record
                quickQuote.db_save()
                # Now let's communicate with other Transana instances if we're in Multi-user mode
                if not TransanaConstants.singleUserVersion:
                    msg = 'Quote %d' % quickQuote.number
                    if TransanaGlobal.chatWindow != None:
                        # Send the "Update Keyword List" message
                        TransanaGlobal.chatWindow.SendMessage("UKL %s" % msg)

                # See if the Keyword visualization needs to be updated.
                dbTree.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Document Keyword Visualization
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Quote', quickQuote.number, quickQuote.source_document_num))

                # Unlock the record
                quickQuote.unlock_record()
            # Handle "RecordLockedError" exception
            except TransanaExceptions.RecordLockedError, e:
                TransanaExceptions.ReportRecordLockedException(_("Quote"), quickQuote.id, e)
            # Handle "SaveError" exception
            except TransanaExceptions.SaveError:
                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
            # Handle other exceptions
            except:
                errordlg = Dialogs.ErrorDialog(None, "%s %s" % (sys.exc_info()[0], sys.exc_info()[1]))
                errordlg.ShowModal()
                errordlg.Destroy()
        # If there is NO duplicate Quick Quote record ...
        else:
            # Determine the next Quick Quote Name
            # Establish the base Quote Name.  Limit the size of the Document Name so we don't overflow
            # the Quote ID field length in any language.
            # Start with the i18n version of "Quick Quote"
            if 'unicode' in wx.PlatformInfo:
                baseName = unicode(_('Quick Quote'), 'utf8') + ' '
            else:
                baseName = _('Quick Quote') + ' '
            # If multi-user ...
            if not TransanaConstants.singleUserVersion:
                # ... then pre-pend the username
                baseName = '- ' + TransanaGlobal.userName + ' - ' + baseName
            # pre-pend as much of the Document name as will fit.
            baseName = sourceDocument.id[:94 - len(baseName)] + ' ' + baseName
            # Start numbering at 1
            baseNum = 1
            # Get a list of all QuickQuotes
            quoteList = DBInterface.list_of_quotes_by_collectionnum(collectNum)
            # Iterate through the list
            for quoteItem in quoteList:
                # Isolate the Quote Name
                quoteName = quoteItem[1]
                # We can ignore errors in converting quote names to integers, but we have to trap it.
                try:
                    # make sure the Base Name matches the comparison Quick Quote name.
                    if baseName == quoteName[:len(baseName)]:
                        # Get the integer portion of the Quote Name that follows the base Quote Name
                        quoteNum = int(quoteName[len(baseName):])
                        # We're looking for the largest value.
                        if quoteNum >= baseNum:
                            # The base number value should always be 1 larger than the largest used value.
                            baseNum = quoteNum + 1
                except:
                    pass

            # Create a Quote Object and populate it with the proper data
            quickQuote = Quote.Quote()
            quickQuote.id = baseName + str(baseNum)
            quickQuote.collection_num = collectNum
            quickQuote.source_document_num = quoteData.documentNum
            quickQuote.start_char = quoteData.startChar
            quickQuote.end_char = quoteData.endChar
            quickQuote.sort_order = DBInterface.getMaxSortOrder(collectNum) + 1
            # Assign the Quote Text
            quickQuote.text = quoteData.text

            # Add the Document Keywords as default Quote Keywords
            quickQuote.keyword_list = sourceDocument.keyword_list
            # Add the keyword that initiated the Quick Quote
            quickQuote.add_keyword(kwg, kw)
            # If there are extra keywords ...
            for (kwg2, kw2) in extraKeywords:
                # ... add them to the Quick Quote
                quickQuote.add_keyword(kwg2, kw2)

            # If we're in multi-user mode ...
            if not TransanaConstants.singleUserVersion:
                if 'unicode' in wx.PlatformInfo:
                    data = (unicode(_("Transana Users"), 'utf8'), TransanaGlobal.userName)
                else:
                    data = (_("Transana Users"), TransanaGlobal.userName)
                # ... determine if the Username is already a Keyword
                if DBInterface.check_username_as_keyword():
                    # Add the new Keyword to the data tree
                    dbTree.add_Node('KeywordNode', (_('Keywords'), _("Transana Users"), TransanaGlobal.userName), 0, _("Transana Users"), expandNode=False)
                    # Inform the Message Server of the added Keyword
                    if TransanaGlobal.chatWindow != None:
                        msg = "AK %s >|< %s"
                        if 'unicode' in wx.PlatformInfo:
                            msg = unicode(msg, 'utf8')
                        TransanaGlobal.chatWindow.SendMessage(msg % data)
                # Add the keyword to the Quick Quote
                quickQuote.add_keyword(data[0], data[1])

            # If we're in Demo mode, an exception can be raised here if the Quote Limit is exceeded.
            # We need to trap that.
            try:
                # Save the Quick Quote
                quickQuote.db_save()
                # Add the QuotePositions to the open Document
                dbTree.parent.ControlObject.AddQuoteToOpenDocument(quickQuote)

                # We need to add the Quote to the database tree.
                # Build the node data needed to add the Quote.
                nodeData = (_('Collections'), collectName, quickQuote.id)
                # Add the new Collection to the data tree
                dbTree.add_Node('QuoteNode', nodeData, quickQuote.number, quickQuote.collection_num, sortOrder=quickQuote.sort_order, expandNode=False)
                # Inform the Message Server of the added Quote
                if not TransanaConstants.singleUserVersion:
                    msg = "AQ %s >|< %s"
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage(msg % (collectName, quickQuote.id))

                # See if the Keyword visualization needs to be updated.
                dbTree.parent.ControlObject.UpdateKeywordVisualization()
                # Even if this computer doesn't need to update the keyword visualization others, might need to.
                if not TransanaConstants.singleUserVersion:
                    # We need to update the Document Keyword Visualization
                    if TransanaGlobal.chatWindow != None:
                        TransanaGlobal.chatWindow.SendMessage("UKV %s %s %s" % ('Document', quickQuote.number, quickQuote.source_document_num))
            # Process the SaveError.  This should ONLY occur in Demo Mode if the Quote Limit is exceeded.
            except TransanaExceptions.SaveError:
                # Remove the part of the error message that refers to clicking "Cancel" in the "Add Quote" dialog box!
                prompt = sys.exc_info()[1].reason.split('\n')[0]

                # Display the Error Message
                errordlg = Dialogs.ErrorDialog(None, prompt)
                errordlg.ShowModal()
                errordlg.Destroy()

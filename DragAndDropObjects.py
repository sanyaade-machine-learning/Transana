# Copyright (C) 2003 - 2006 The Board of Regents of the University of Wisconsin System 
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
     method to allow code resuse by several different methods related to dropping and pasting clips.  """

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "DragAndDropObjects DEBUG is ON!!"


import wx                           # Import wxPython
import cPickle                      # use Python's fast cPickle tool instead of the regular Pickle
import sys                          # import Python's sys module

import DBInterface                  # Import Transana's Database Interface
import Series                       # Import the Transana Series object
import Episode                      # Import the Transana Episode Object
import Collection                   # Import the Transana Collection Object
import Clip                         # Import the Transana Clip Object
import Transcript                   # Import the Transana Transcript Object
import ClipPropertiesForm           # Import the Transana Clip Properties Form for adding Clips on Transcript Text Drop
import Keyword                      # Import the Transana Keyword Object
import DatabaseTreeTab              # Import the Transana Database Tree Tab Object (for setting _NodeData in manipulating the tree)
import Misc                         # Import the Transana Miscellaneous routines
import Dialogs                      # Import the Transana Dialog Boxes
import TransanaConstants            # Import the Transana Constants
import TransanaGlobal               # Import Transana's Globals
import TransanaExceptions           # Import Transana's Exceptions


def DragDropEvaluation(source, destination):
   """ This boolean function indicates whether the source tree node can legally be dropped (or pasted) on the destination
       tree node.  This function is encapsulated because it needs to be called from several different locations
       during the Drag-and-Drop process, including the DropSource's GiveFeedback() Method and the DropTarget's
       OnData() Method, as well as the DBTree's OnRightClick() to enable or disable the "Paste" option. """
   # Return True if the drop is legal, false if it is not.
   # To be legal, we must have a legitimate source and be on a legitimate drop target.
   # If the source is the Database Tree Tab (nodetype = DataTreeDragDropData), then we compare
   # the nodetypes for the source and destination nodes to see if the pairing is compatible.
   # Next, either the record numbers or the nodetype must be different, so you can't drop a node on itself.
   if (source != None) and \
      (destination != None) and \
      (type(source) != type(ClipDragDropData())) and \
      ((source.nodetype == 'CollectionNode'       and destination.nodetype == 'CollectionNode') or \
       (source.nodetype == 'ClipNode'             and destination.nodetype == 'CollectionNode') or \
       (source.nodetype == 'ClipNode'             and destination.nodetype == 'ClipNode') or \
       (source.nodetype == 'ClipNode'             and destination.nodetype == 'KeywordNode') or \
       (source.nodetype == 'KeywordNode'          and destination.nodetype == 'SeriesNode') or \
       (source.nodetype == 'KeywordNode'          and destination.nodetype == 'EpisodeNode') or \
       (source.nodetype == 'KeywordNode'          and destination.nodetype == 'CollectionNode') or \
       (source.nodetype == 'KeywordNode'          and destination.nodetype == 'ClipNode') or \
       (source.nodetype == 'KeywordNode'          and destination.nodetype == 'KeywordGroupNode') or \
       (source.nodetype == 'SearchCollectionNode' and destination.nodetype == 'SearchCollectionNode') or \
       (source.nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchCollectionNode') or \
       (source.nodetype == 'SearchClipNode'       and destination.nodetype == 'SearchClipNode')) and \
      ((source.recNum != destination.recNum) or (source.nodetype != destination.nodetype)):
       return True
   # If we have a Clip (type == ClipDragDropData), then we can drop it on a Collection or a Clip only.
   elif (source != None) and \
        (destination != None) and \
        (type(source) == type(ClipDragDropData())) and \
	((destination.nodetype == 'CollectionNode') or \
	 (destination.nodetype == 'ClipNode')):
       return True
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
      self.nodeList = nodeList   # The Source Node's nodeList (for SearchCollectionNode and SearchClipNode Cut and Paste)
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
#      print "window (x,y) = (%d,%d)" % (windowx, windowy)
#      print "(Left=%d, Right=%d, Top=%d, Bot=%d" % (transLeft, transRight, transTop, transBot)

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

          # print "DataTreeDropSource.GiveFeedback: (%s, %s), (%s, %s)" % (x, y, w, h)

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

      # Create a Composite Data Object
      self.doc = wx.DataObjectComposite()
      # Add the Tree Node Data Object
      self.doc.Add(self.sourceNodeData)
      # Add the Clip Data Object
      self.doc.Add(self.clipData)
      
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
       # Replace teh old Clip Creation Data Object with the new empty one
       self.clipData.SetData(pickledTempData)

   def OnEnter(self, x, y, dragResult):
      # Just allow the normal wxDragResult to pass through here
      return dragResult

   def OnLeave(self):
      pass

   def OnDrop(self, x, y):
      # Process the "Drop" event
      
      # print "Drop:  x=%s, y=%s" % (x, y)

      # If you drop off the Database Tree, you get an exception here
      try:
          # Use the tree's HitTest method to find out about the potential drop target for the current mouse position
          (self.dropNode, flag) = self.tree.HitTest((x, y))
          # Remember the Drop Location for later Processing (in OnData())
          self.dropData = self.tree.GetPyData(self.dropNode)

          # print "Dropped on %s." % self.dropData

          # We don't yet have enough information to veto the drop, so return TRUE to indicate
          # that we should proceed to the OnData method
          return True
      except:
          # If an exception is raised, Veto the drop as there is no Drop Target.
          return False

   def OnData(self, x, y, dragResult):
      # once OnDrop returns TRUE, this method is automatically called.
      
      # print "OnData %s, %s, %s" % (x, y, dragResult)

      # Let's get the data being dropped so we can do some processing logic
      if self.GetData():
         # First, extract the actual data passed in by the DataTreeDropSource, which used cPickle to pack it.

         try:
             # Try to unPickle the Tree Node Data Object.  If the first Drag is for Clip Creation, this
             # will raise an exception.  If there is a good Tree Node Data Object being dragged, or if
             # one from a previous drag has been Cleared, this will be successful.
             sourceData = cPickle.loads(self.sourceNodeData.GetData())

             # print "Drop Data = '%s' dropped onto %s" % (sourceData, self.dropData)
             # print "sourceData is of type ", type(sourceData)

             # if type(sourceData) == type(DataTreeDragDropData()):
             #     print "type is DataTreeDragDropData"
             # elif type(sourceData) == type(ClipDragDropData()):
             #     print "type is ClipDragDropData"
             # else:
             #     print 'type is unknown'


             # If a previous drag of a Tree Node Data Object has been cleared, the sourceData.nodetype
             # will be "Unknown", which indicated that the current Drag is NOT a node from the Database
             # Tree Tab, and therefore should be processed elsewhere.  If it is NOT "Unknown", we should
             # process it here.  The Type comparison was added to get this working on the Mac.
             if (type(sourceData) == type(DataTreeDragDropData())) and \
                (sourceData.nodetype != 'Unknown'):
                 # This line compares the data being dragged (sourceData) to the drop site determined in OnDrop and
                 # passed here as self.dropData.  
                 if DragDropEvaluation(sourceData, self.dropData):
                    # If we meet the criteria, we process the drop.  We do that here because we have full
                    # knowledge of the Dragged Data and the Drop Target's data here and nowhere else.
                    # Determine if we're copying or moving data.  (In some instances, the 'action' is ignored.)
                    if dragResult == wx.DragCopy:
                       ProcessPasteDrop(self.tree, sourceData, self.dropNode, 'Copy')
                    elif dragResult == wx.DragMove:
                       ProcessPasteDrop(self.tree, sourceData, self.dropNode, 'Move')
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
             if '__WXMAC__' in wx.PlatformInfo:
                 clipData = sourceData
             else:
                 clipData = cPickle.loads(self.clipData.GetData())

             # print "self.clipData =\n", clipData
             # print "clipData is of type ", type(clipData)

             # if type(clipData) == type(DataTreeDragDropData()):
             #     print "type 2 is DataTreeDragDropData"
             # elif type(clipData) == type(ClipDragDropData()):
             #     print "type 2 is ClipDragDropData"
             # else:
             #     print 'type 2 is unknown'

             # if '__WXMAC__' in wx.PlatformInfo:
             #     print "Second Node Check skipped for Mac"
             # else:
             #     if sourceData.nodetype == 'Unknown':
             #        print "Drop is NOT a Node"
             #     else:
             #        print "Drop is a Node"
             # if clipData.transcriptNum == 0:
             #    print "Drop is NOT a Clip"
             # else:
             #    print "Drop is a Clip"

             # print

             # See if the Drop Target is the correct Node Type.  The type comparison was added to get this working on the Mac.
             if (type(clipData) == type(ClipDragDropData())) and \
                ((self.dropData.nodetype == 'CollectionNode') or \
                 (self.dropData.nodetype == 'ClipNode')):

                 # If a previous drag of a Clip Creation Data Object has been cleared, the clipData.transcriptNum
                 # will be "0", which indicated that the current Drag is NOT a Clip Creation Data Object, 
                 # and therefore should be processed elsewhere.  If it is NOT "0", we should process it here.
                 if clipData.transcriptNum != 0:
                     
                     CreateClip(clipData, self.dropData, self.tree, self.dropNode)
                     
                         # Once the drop is done or rejected, we must clear the Clip Creation data out of the DropTarget.
                         # If we don't, this data will still be there if a Tree Node drag occurs, and there is no way in that
                         # circumstance to know which of the dragged objects to process!  Clearing avoids that problem.
                     self.ClearClipData()
                     
             else:
                # If the Drop target is not valid, we prevent the drop process by altering the wxDropResult (dragResult)
                dragResult = wx.DragNone
             # Reset the cursor, regardless of whether the drop succeeded or failed.
             self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))
                 
         except:
             # Reset the cursor, regardless of whether the drop succeeded or failed.
             self.tree.SetCursor(wx.StockCursor(wx.CURSOR_ARROW))

             import sys
             (exType, exValue) =  sys.exc_info()[:2]
         
             # If an expection occurs here, it's no big deal.  Forget about it.
             pass
            
      # Returning this value allows us to confirm or veto the drop request
      return dragResult


class ClipDragDropData(object):
   """ This object contains all the data that needs to be transferred in order to create a Clip
       from a selection in a Transcript. """

   def __init__(self, transcriptNum=0, episodeNum=0, clipStart=0, clipStop=0, text=''):
      """ ClipDragDropData Objects require the following parameters:
          transcriptNum    The Transcript Number of the originating Transcript
          episodeNum       The Episode the originating Transcript is attached to
          clipStart        The starting Time Code for the Clip
          clipStop         the ending Time Code for the Clip
          text             the Text for the Clip, in RTF format. """
      self.transcriptNum = transcriptNum
      self.episodeNum = episodeNum
      self.clipStart = clipStart
      self.clipStop = clipStop
      self.text = text

   def __repr__(self):
      str = 'ClipDragDropData Object:\n'
      str = str + 'transcriptNum = %s\n' % self.transcriptNum
      str = str + 'episodeNum = %s\n' % self.episodeNum
      str = str + 'clipStart = %s\n' % Misc.time_in_ms_to_str(self.clipStart)
      str = str + 'clipStop = %s\n' % Misc.time_in_ms_to_str(self.clipStop)
      str = str + 'text = %s\n\n' % self.text
      return str

      
def CreateClip(clipData, dropData, tree, dropNode):
    """ This method handles the creation of a Clip Object in the Transana Database """
    # Create a new Clip Object
    tempClip = Clip.Clip()
    # We need to know if the Clip is coming from an Episode or another Clip.
    # We can determine that by looking at the transcript passed in the ClipData
    tempTranscript = Transcript.Transcript(clipData.transcriptNum)
    # Get the Episode Number from the clipData Object
    tempClip.episode_num = clipData.episodeNum
    # If we are working from an Episode Transcript ...
    if tempTranscript.clip_num == 0:
        # Get the Transcript Number from the clipData Object
        tempClip.transcript_num = clipData.transcriptNum
    # If we are working from a Clip Transcript ...
    else:
        sourceClip = Clip.Clip(tempTranscript.clip_num)
        tempClip.transcript_num = sourceClip.transcript_num
    # Get the Clip Start Time from the clipData Object
    tempClip.clip_start = clipData.clipStart
    # Get the Clip Stop Time from the clipData Object
    tempClip.clip_stop = clipData.clipStop
    # Get the Clip Transcript from the clipData Object
    tempClip.text = clipData.text

    # If the Clip Creation Object is dropped on a Collection ...
    if dropData.nodetype == 'CollectionNode':
        # ... get the Clip's Collection Number from the Drop Node ...
        tempClip.collection_num = dropData.recNum
        # ... and the Clip's Collection Name from the Drop Node.
        tempClip.collection_id = tree.GetItemText(dropNode)
        # Remember the Collection Node which should be the parent for the new Clip Node to be created later.
        collectionNode = dropNode
    # If the Clip Creation Object is dropped on a Clip ...
    elif dropData.nodetype == 'ClipNode':
        # ... get the Clip's Collection Number from the Drop Node's Parent ...
        tempClip.collection_num = dropData.parent
        # ... and the Clip's Collection Name from the Drop Node's Parent.
        tempClip.collection_id = tree.GetItemText(tree.GetItemParent(dropNode))
        # Remember the Collection Node which should be the parent for the new Clip Node to be created later.
        collectionNode = tree.GetItemParent(dropNode)
    # Load the Episode that is connected to the Clip's Originating Transcript
    tempEpisode = Episode.Episode(tempClip.episode_num)
    # The Clip's Media Filename comes from the Episode Record
    tempClip.media_filename = tempEpisode.media_filename
    # Load the parent Collection
    tempCollection = Collection.Collection(tempClip.collection_num)
    try:
        # Lock the parent Collection, to prevent it from being deleted out from under the add.
        tempCollection.lock_record()
        collectionLocked = True
    # Handle the exception if the record is already locked by someone else
    except RecordLockedError, c:
        # If we can't get a lock on the Collection, it's really not that big a deal.  We only try to get it
        # to prevent someone from deleting it out from under us, which is pretty unlikely.  But we should 
        # still be able to add Clips even if someone else is editing the Collection properties.
        collectionLocked = False
    # Create the Clip Properties Dialog Box to Add a Clip
    dlg = ClipPropertiesForm.AddClipDialog(None, -1, tempClip)
    # Set the "continue" flag to True (used to redisplay the dialog if an exception is raised)
    contin = True
    # While the "continue" flag is True ...
    while contin:
        # Display the Clip Properties Dialog Box and get the data from the user
        if dlg.get_input() != None:
            # Use "try", as exceptions could occur
            try:
                # See if the Clip Name already exists in the Destination Collection
                (dupResult, newClipName) = CheckForDuplicateClipName(tempClip.id, tree, dropNode)

                # If a Duplicate Clip Name is found and the error situation not resolved, show an Error Message
                if dupResult:
                    if 'unicode' in wx.PlatformInfo:
                        # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                        prompt = unicode(_('Clip Creation cancelled for Clip "%s".  Duplicate Clip Name Error.'), 'utf8') % tempClip.id
                    else:
                        prompt = _('Clip Creation cancelled for Clip "%s".  Duplicate Clip Name Error.') % tempClip.id
                    errordlg = Dialogs.ErrorDialog(None, prompt)
                    errordlg.ShowModal()
                    errordlg.Destroy()
                    # Unlock the parent collection
                    if collectionLocked:
                        tempCollection.unlock_record()
                    # If the user cancels Clip Creation, we don't need to continue any more.
                    contin = False
                else:
                    # If the Name was changed, reflect that in the Clip Object
                    tempClip.id = newClipName
                    # See if we're dropping on a Collection Node ...
                    if dropData.nodetype == 'CollectionNode':
                        tempClip.sort_order = tree.GetChildrenCount(dropNode, False) + 1
                    # Try to save the data from the form
                    tempClip.db_save()
                    tempCollection = Collection.Collection(tempClip.collection_num)
                    nodeData = (_('Collections'),) + tempCollection.GetNodeData() + (tempClip.id,)

                    # See if we're dropping on a Clip Node ...
                    if dropData.nodetype == 'ClipNode':
                        # Add the new Collection to the data tree
                        tree.add_Node('ClipNode', nodeData, tempClip.number, tempClip.collection_num, True, dropNode)
                    else:
                        # Add the new Clip to the data tree
                        tree.add_Node('ClipNode', nodeData, tempClip.number, tempClip.collection_num)

                    # Now let's communicate with other Transana instances if we're in Multi-user mode
                    if not TransanaConstants.singleUserVersion:
                        msg = "ACl %s"
                        data = (nodeData[1],)

                        for nd in nodeData[2:]:
                            msg += " >|< %s"
                            data += (nd, )
                        if TransanaGlobal.chatWindow != None:
                            TransanaGlobal.chatWindow.SendMessage(msg % data)

                    # See if we're dropping on a Clip Node ...
                    if dropData.nodetype == 'ClipNode':
                        # ... and if so, change the Sort Order of the clips
                        ChangeClipOrder(tree, dropNode, tempClip, tempCollection)

                    # Unlock the parent collection
                    if collectionLocked:
                        tempCollection.unlock_record()
                    # If we do all this, we don't need to continue any more.
                    contin = False
            # Handle "SaveError" exception
            except TransanaExceptions.SaveError:
                # Display the Error Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
                errordlg.ShowModal()
                errordlg.Destroy()
            # Handle other exceptions
            except:
                # Display the Exception Message, allow "continue" flag to remain true
                errordlg = Dialogs.ErrorDialog(None, "%s" % (sys.exc_info()[:2], ))
                errordlg.ShowModal()
                errordlg.Destroy()

                import traceback
                print traceback.print_exc()
                
        # If the user pressed Cancel ...
        else:
            # Unlock the parent collection
            if collectionLocked:
                tempCollection.unlock_record()
            # ... then we don't need to continue any more.
            contin = False
    dlg.Destroy()


def DropKeyword(parent, sourceData, targetType, targetName, targetRecNum, targetParent):
    """Drop a Keyword onto an Object.  sourceData is from the Keyword.  The
    targetType is one of 'Series', 'Episode', 'Collection', or 'Clip'.
    targetParent is only used for collections."""
    
    # Trying to deal with a Mac issue temporarily.  Dragging within the Transcript sometimes triggers this method when it shouldn't.  Can't find the cause.
    if "__WXMAC__" in wx.PlatformInfo and \
       (type(sourceData) == type(ClipDragDropData())):
           return
    
    if targetType == 'Series':
        # Get user confirmation of the Keyword Add request
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to add Keyword "%s:%s" to all the Episodes in\nSeries "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
        else:
            prompt = _('Do you want to add Keyword "%s:%s" to all the Episodes in\nSeries "%s"?') % (sourceData.parent, sourceData.text, targetName)
        dlg = wx.MessageDialog(parent,  prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
        
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_NO:
            return
        # If confirmed, copy the Keyword to all Episodes in the Series
        # print "Keyword %s:%s to be dropped on Series %s" % (sourceData.parent, sourceData.text, treeCtrl.GetItemText(destNode))
        # First, let's load the Series Record
        tempSeries = Series.Series(targetRecNum)
        # Lock the Series Record, just to be on the safe side (Is this necessary??  I don't think so, but maybe that can confirm that all episodes are available.)
        tempSeries.lock_record()
        # Now get a list of all Episodes in the Series and iterate through them
        for tempEpisodeNum, tempEpisodeID, tempSeriesNum in DBInterface.list_of_episodes_for_series(tempSeries.id):
            # Load the Episode Record
            tempEpisode = Episode.Episode(num=tempEpisodeNum)
            # Lock the Episode Record
            tempEpisode.lock_record()
            # Add the Keyword to the Episode
            tempEpisode.add_keyword(sourceData.parent, sourceData.text)
            # Save the Episode
            tempEpisode.db_save()
            # Unlock the Episode Record
            tempEpisode.unlock_record()
        # Unlock the Series Record
        tempSeries.unlock_record()   
    
    elif targetType == 'Episode':
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to add Keyword "%s:%s" to\nEpisode "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
        else:
            prompt = _('Do you want to add Keyword "%s:%s" to\nEpisode "%s"?') % (sourceData.parent, sourceData.text, targetName)
        dlg = wx.MessageDialog(parent, prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_NO:
            return
        # If confirmed, copy the Keyword to the Episodes
        # print "Keyword %s:%s to be dropped on Episode %s" % (sourceData.parent, sourceData.text, treeCtrl.GetItemText(destNode))
        # Load the Episode Record
        tempEpisode = Episode.Episode(num=targetRecNum)
        # Lock the Episode Record
        tempEpisode.lock_record()
        # Add the keyword to the Episode
        tempEpisode.add_keyword(sourceData.parent, sourceData.text)
        # Save the Episode
        tempEpisode.db_save()
        # Unlock the Episode
        tempEpisode.unlock_record()
    
    elif targetType == 'Collection':
        # Get user confirmation of the Keyword Add request
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to add Keyword "%s:%s" to all the Clips in\nCollection "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
        else:
            prompt = _('Do you want to add Keyword "%s:%s" to all the Clips in\nCollection "%s"?') % (sourceData.parent, sourceData.text, targetName)
        dlg = wx.MessageDialog(parent,  prompt , _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_NO:
            return
        # If confirmed, copy the Keyword to all Clips in the Collection
        # print "Keyword %s:%s to be dropped on Collection %s" % (sourceData.parent, sourceData.text, treeCtrl.GetItemText(destNode))
        # First, load the Collection
        tempCollection = Collection.Collection(targetRecNum, targetParent)
        # Lock the Collection Record, just to be on the safe side (Is this necessary??  I don't think so, but maybe that can confirm that all Clips are available.)
        tempCollection.lock_record()
        # Now load a list of all the Clips in the Collection and iterate through them
        for tempClipNum, tempClipID, tempCollectNum in DBInterface.list_of_clips_by_collection(tempCollection.id, tempCollection.parent):
            # Load the Clip
            tempClip = Clip.Clip(id_or_num=tempClipNum)
            # Lock the Clip
            tempClip.lock_record()
            # Add the Keyword to the Clip
            tempClip.add_keyword(sourceData.parent, sourceData.text)
            # Save the Clip
            tempClip.db_save()
            # Unlock the Clip
            tempClip.unlock_record()
        # Unlock the Collection Record
        tempCollection.unlock_record()
    
    elif targetType == 'Clip':
        # Get user confirmation of the Keyword Add request
        if 'unicode' in wx.PlatformInfo:
            # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
            prompt = unicode(_('Do you want to add Keyword "%s:%s" to\nClip "%s"?'), 'utf8') % (sourceData.parent, sourceData.text, targetName)
        else:
            prompt = _('Do you want to add Keyword "%s:%s" to\nClip "%s"?') % (sourceData.parent, sourceData.text, targetName)
        dlg = wx.MessageDialog(parent,  prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_NO:
            return
        # If confirmed, copy the Keyword to the Clip
        # First, load the Clip
        tempClip = Clip.Clip(id_or_num=targetRecNum)
        # Lock the Clip Record
        tempClip.lock_record()
        # Add the Keyword to the Clip Record
        tempClip.add_keyword(sourceData.parent, sourceData.text)

        try:
            # Save the Clip Record
            tempClip.db_save()

        # Handle "SaveError" exception
        except SaveError:
            # Display the Error Message, allow "continue" flag to remain true
            errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
            errordlg.ShowModal()
            errordlg.Destroy()
        # Handle other exceptions
        except:
            # Display the Exception Message, allow "continue" flag to remain true
            errordlg = Dialogs.ErrorDialog(None, "%s" % (sys.exc_info()[:2]))
            errordlg.ShowModal()
            errordlg.Destroy()

        # Unlock the Clip Record
        tempClip.unlock_record()

def ProcessPasteDrop(treeCtrl, sourceData, destNode, action):
   """ This method processes a "Paste" or "Drop" request for the Transana Database Tree.
       Parameters are:
         treeCtrl   -- the wxTreeCtrl where the Paste or Drop is occurring (the DBTree)
         sourceData -- the DATA associated with the Cut/Copy or Drag, the _NodeData or the DataTreeDragDropData Object
         destNode   -- the actual Tree Node selected for Drop or Paste
         action     -- a string of "Copy" or "Move", indicating whether a Copy or Cut/Move has been requested.
                       (This value is ignored in some instances where "Move" has no meaning.  """

   # Since we get the actual destination node as a parameter, let's first extract the Node Data for the Destination
   destNodeData = treeCtrl.GetPyData(destNode)
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

   # Drop a Collection on a Collection (Copy or Move all Clips in a Collection)
   if (sourceData.nodetype == 'CollectionNode' and destNodeData.nodetype == 'CollectionNode'):
      # Load the Source Collection
      sourceCollection = Collection.Collection(sourceData.recNum, sourceData.parent)
      # We need a list of all the clips in the Source Collection
      clipList = DBInterface.list_of_clips_by_collection(sourceCollection.id, sourceCollection.parent)
      # Load the Destination Collection
      destCollection = Collection.Collection(destNodeData.recNum, destNodeData.parent)

      # Get user confirmation of the Collection Copy request
      if 'unicode' in wx.PlatformInfo:
         # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
         prompt = unicode(_('Do you want to %s all Clips from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceCollection.id, destCollection.id)
      else:
         prompt = _('Do you want to %s all Clips from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceCollection.id, destCollection.id)
      dlg = wx.MessageDialog(treeCtrl,  prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
      result = dlg.ShowModal()
      dlg.Destroy()
      if result == wx.ID_YES:
         # If confirmed, iterate through the list of Clips
         for clip in clipList:
            # Load the next Clip from the list
            tempClip = Clip.Clip(id_or_num=clip[0])
            # Copy or Move the Clip to the Destination Collection
            CopyMoveClip(treeCtrl, destNode, tempClip, sourceCollection, destCollection, action)
               
   # Drop a Clip on a Collection (Copy or Move a Clip)
   elif (sourceData.nodetype == 'ClipNode' and destNodeData.nodetype == 'CollectionNode'):
      # Load the Source Clip
      sourceClip = Clip.Clip(id_or_num=sourceData.recNum)
      # Load the Source Collection
      sourceCollection = Collection.Collection(sourceData.parent)
      # Load the Destination Collection
      destCollection = Collection.Collection(destNodeData.recNum, destNodeData.parent)

      # Get user confirmation of the Clip Copy/Move request
      if 'unicode' in wx.PlatformInfo:
         # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
         prompt = unicode(_('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
      else:
         prompt = _('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
      dlg = wx.MessageDialog(treeCtrl, prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
      result = dlg.ShowModal()
      dlg.Destroy()
      if result == wx.ID_YES:
         # Copy or Move the Clip to the Destination Collection
         CopyMoveClip(treeCtrl, destNode, sourceClip, sourceCollection, destCollection, action)
               
   # Drop a Clip on a Clip (Alter SortOrder, Copy or Move a Clip to a particular place in the SortOrder)
   elif (sourceData.nodetype == 'ClipNode' and destNodeData.nodetype == 'ClipNode'):
      # Load the Source Clip
      sourceClip = Clip.Clip(id_or_num=sourceData.recNum)
      # Load the Source Collection
      sourceCollection = Collection.Collection(sourceData.parent)
      # Load the Destination (or target) Clip
      destClip = Clip.Clip(destNodeData.recNum)
      # Load the Destination (or target) Collection
      destCollection = Collection.Collection(destClip.collection_num)

      # See if we are in the SAME Collection, and therefore just changing Sort Order
      if sourceCollection.number == destCollection.number:
         # If so, change the Sort Order as requested
         ChangeClipOrder(treeCtrl, destNode, sourceClip, sourceCollection)
         # NOTE:  We can't just use the insertPos parameter of add_Node, as we need to have the Sort Order set in the
         #        Clip Object as well.

      # If not, we are copying/moving a Clip to a place in the SortOrder
      else:
         # Get user confirmation of the Clip Copy/Move request
         if 'unicode' in wx.PlatformInfo:
             # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
             prompt = unicode(_('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?'), 'utf8') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
         else:
             prompt = _('Do you want to %s Clip "%s" from\nCollection "%s" to\nCollection "%s"?') % (copyMovePrompt, sourceClip.id, sourceCollection.id, destCollection.id)
         dlg = wx.MessageDialog(treeCtrl, prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
         result = dlg.ShowModal()
         dlg.Destroy()
         if result == wx.ID_YES:
            # Copy or Move the Clip to the Destination Collection
            # If confirmed, copy the Source Clip to the Destination Collection.  CopyClip will place the clip at
            # end of the list of clips.
            # We need to work with the COPY of the clip instead of the original from here on, so we get that
            # value from CopyClip.
            tempClip = CopyMoveClip(treeCtrl, destNode, sourceClip, sourceCollection, destCollection, action)
            # If the Copy/Move is cancelled, tempClip will be None
            if tempClip != None:
                # Now change the order of the clips
                ChangeClipOrder(treeCtrl, destNode, tempClip, destCollection)

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
      dlg = wx.MessageDialog(treeCtrl, prompt, "Transana Confirmation", style=wx.YES_NO | wx.ICON_QUESTION)
      result = dlg.ShowModal()
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
            
            if DEBUG:
               print 'Message to send = "AKE %d >|< %s >|< %s >|< %s"' % (sourceData.recNum, kwg, kw, sourceData.text)
               
            if TransanaGlobal.chatWindow != None:
               TransanaGlobal.chatWindow.SendMessage("AKE %d >|< %s >|< %s >|< %s" % (sourceData.recNum, kwg, kw, sourceData.text))

   # Drop a Keyword on a Series
   elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'SeriesNode'):
      DropKeyword(treeCtrl, sourceData, 'Series', \
              treeCtrl.GetItemText(destNode), destNodeData.recNum, 0)
   
   # Drop a Keyword on an Episode
   elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'EpisodeNode'):
      DropKeyword(treeCtrl, sourceData, 'Episode', \
              treeCtrl.GetItemText(destNode), destNodeData.recNum, 0)
   
      # Drop a Keyword on a Collection
   elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'CollectionNode'):
      DropKeyword(treeCtrl, sourceData, 'Collection', \
              treeCtrl.GetItemText(destNode), destNodeData.recNum, destNodeData.parent)

   # Drop a Keyword on a Clip
   elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'ClipNode'):
      DropKeyword(treeCtrl, sourceData, 'Clip', \
              treeCtrl.GetItemText(destNode), destNodeData.recNum, 0)

   # Drop a Keyword on a Keyword Group (Copy or Move a Keyword)
   elif (sourceData.nodetype == 'KeywordNode' and destNodeData.nodetype == 'KeywordGroupNode'):
      # Get user confirmation of the Keyword Copy request
      if 'unicode' in wx.PlatformInfo:
         # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
         prompt = unicode(_('Do you want to %s Keyword "%s" from\nKeyword Group "%s" to\nKeyword Group "%s"?'), 'utf8') % (copyMovePrompt, sourceData.text, sourceData.parent, treeCtrl.GetItemText(destNode))
      else:
         prompt = _('Do you want to %s Keyword "%s" from\nKeyword Group "%s" to\nKeyword Group "%s"?') % (copyMovePrompt, sourceData.text, sourceData.parent, treeCtrl.GetItemText(destNode))
      dlg = wx.MessageDialog(treeCtrl, prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
      result = dlg.ShowModal()
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
         elif action == 'Move':
            # If confirmed, move the Keyword to the Clip.
            # Load the appropriate Keyword Object
            tempKeyword = Keyword.Keyword(sourceData.parent, sourceData.text)
            # Lock the Keyword Record
            tempKeyword.lock_record()

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
                if DEBUG:
                   print 'Message to send = "AK %s >|< %s"' % (treeCtrl.GetItemText(destNode), sourceData.text)
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
                        
                        if DEBUG:
                           print 'Message to send = "AKE %d >|< %s >|< %s >|< %s"' % (nodeData.recNum, treeCtrl.GetItemText(destNode), sourceData.text, nodeName)
                           
                        if TransanaGlobal.chatWindow != None:
                           TransanaGlobal.chatWindow.SendMessage("AKE %d >|< %s >|< %s >|< %s" % (nodeData.recNum, treeCtrl.GetItemText(destNode), sourceData.text, nodeName))
         except TransanaExceptions.SaveError:
             # Display the Error Message, allow "continue" flag to remain true
             errordlg = Dialogs.ErrorDialog(None, sys.exc_info()[1].reason)
             errordlg.ShowModal()
             errordlg.Destroy()


   # Drop a SearchCollection on a SearchCollection (Copy or Move all SearchClips in a SearchCollection)
   elif (sourceData.nodetype == 'SearchCollectionNode' and destNodeData.nodetype == 'SearchCollectionNode'):
      # NOTE:  SearchClips don't exist in the database.  Therefore, to copy or move them,
      #        all we need to do is manipulate Database Tree Nodes

      # Get user confirmation of the Collection Copy request.
      # First, set the names for use in the prompt.
      sourceCollectionId = sourceData.text
      destCollectionId = treeCtrl.GetItemText(destNode)
      # Create the Dialog Box for the confirmation prompt
      if 'unicode' in wx.PlatformInfo:
         # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
         prompt = unicode(_('Do you want to %s all Search Results Clips from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?'), 'utf8') % (copyMovePrompt, sourceCollectionId, destCollectionId)
      else:
         prompt = _('Do you want to %s all Search Results Clips from\nSearch Results Collection "%s" to\nSearch Results Collection "%s"?') % (copyMovePrompt, sourceCollectionId, destCollectionId)
      dlg = wx.MessageDialog(treeCtrl, prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
      # Show the confirmation prompt Dialog
      result = dlg.ShowModal()
      # Clean up after the confirmation Dialog
      dlg.Destroy()
      # Process the results only if the user confirmed.
      if result == wx.ID_YES:
         # We need to keep a record of items that need to be deleted later.  Initialize a list for that.
         # (We can't just delete records as we go as that screws up iterating through the Collection's child nodes.)
         itemsToDelete = []

         # Let's get all the children of the Source SearchCollection.  However, we don't have direct access to the Source
         # Node in all cases.  (We can get it in Drag and Drop, but not Cut and Paste.  I tried to pass the actual
         # node in the SourceData, but that caused catastrophic program crashes.)
         # This ALWAYS gets the correct Source Node.
         sourceNode = treeCtrl.select_Node(sourceData.nodeList, sourceData.nodetype)
         # Iterating through a wxTreeCtrl requires a "cookie" value.  Initialize it here.
         cookieItem = 0L
         # Get the Source Collection's first Child Node
         (childNode, cookieItem) = treeCtrl.GetFirstChild(sourceNode)
         # Iterate through all the Source Collection's Child Nodes
         while childNode.IsOk():
            # Get the Child Node's Node Data
            childData = treeCtrl.GetPyData(childNode)
            # We should only move Clips, not nested Collections
            if childData.nodetype == 'SearchClipNode':
               # We need to reference the child's Name repeatedly, so let's create a variable for it.
               clipName = treeCtrl.GetItemText(childNode)
               # Check for Duplicate Clip Names, an error condition
               (dupResult, newClipName) = CheckForDuplicateClipName(clipName, treeCtrl, destNode)
               # If a Duplicate Clip Name is found and the error situation not resolved, show an Error Message
               if dupResult:
                  if 'unicode' in wx.PlatformInfo:
                     # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                     prompt = unicode(_('%s cancelled for Clip "%s".  Duplicate Clip Name Error.'), 'utf8') % (copyMovePrompt, clipName)
                  else:
                     prompt = _('%s cancelled for Clip "%s".  Duplicate Clip Name Error.') % (copyMovePrompt, clipName)
                  dlg = Dialogs.ErrorDialog(None, prompt)
                  dlg.ShowModal()
                  dlg.Destroy()
               else:
                  # The user may have given the clip a new Clip Name in CheckForDuplicateClipName.  
                  if newClipName != sourceData.text:
                     # If so, use this new name!
                     clipName = newClipName

                  # What we need to do first is add the appropriate new Node to the Tree.
                  # Let's build the NodeList by starting with the Source Clip text ...
                  nodeList = (clipName,)
                  # ... and then climbing the Destination Node Tree .... 
                  currentNode = destNode
                  # ... until we get to the SearchRootNode
                  while treeCtrl.GetPyData(currentNode).nodetype != 'SearchRootNode':
                     # We add the Item Test to the FRONT of the Node List
                     nodeList = (treeCtrl.GetItemText(currentNode),) + nodeList
                     # and we move up to the node's parent
                     currentNode = treeCtrl.GetItemParent(currentNode)
                  # Now Add the new Node, using the SourceData's Data
                  treeCtrl.add_Node('SearchClipNode', (_('Search'),) + nodeList, childData.recNum, childData.parent, False)
                  # No need to communicate with other Transana Clients here, we're just manipulating Search Results.

                  # If we need to remove the node, the SourceData carries the nodeList to the Collection, but we can't just
                  # delete that because it might have subCollections.  Therefore, add the Clip Name to the CourseCollection
                  # Node List 
                  if action == 'Move':
                     itemsToDelete.append((sourceData.nodeList + (treeCtrl.GetItemText(childNode),), treeCtrl.GetPyData(childNode).nodetype),)

            # If we are not currently looking at the Source Collection's LAST Child ...
            if childNode != treeCtrl.GetLastChild(sourceNode):
               # ... then let's move on to the next Child record.
               (childNode, cookieItem) = treeCtrl.GetNextChild(sourceNode, cookieItem)
            # Otherwise ...
            else:
               # We've seen all there is to see, so stop iterating.
               break

         # Delete any items that have been flagged for deletion.
         for (item, itemNodeType) in itemsToDelete:
            treeCtrl.delete_Node(item, itemNodeType)

         if (action == 'Move') and (len(itemsToDelete) > 0):
             # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
             ClearClipboard()
            
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
      dlg = wx.MessageDialog(treeCtrl, prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
      # Display the confirmation Dialog Box
      result = dlg.ShowModal()
      # Clean up after the confirmation Dialog box
      dlg.Destroy()
      # If the user confirmed, process the request.
      if result == wx.ID_YES:
         # Check for Duplicate Clip Name error
         (dupResult, newClipName) = CheckForDuplicateClipName(sourceData.text, treeCtrl, destNode)
         # If a Duplicate Clip Name exists that is not resolved within CheckForDuplicateClipNames ...
         if dupResult:
            # ... display the error message.
            if 'unicode' in wx.PlatformInfo:
                # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
                prompt = unicode(_('%s cancelled for Clip "%s".  Duplicate Clip Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
            else:
                prompt = _('%s cancelled for Clip "%s".  Duplicate Clip Name Error.') % (copyMovePrompt, sourceData.text)
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
            # Now Add the new Node, using the SourceData's Data
            treeCtrl.add_Node('SearchClipNode', (_('Search'),) + nodeList, sourceData.recNum, sourceData.parent, False)
            # No need to communicate with other Transana Clients here, we're just manipulating Search Results.
            # If we need to remove the node, the SourceData carries the nodeList we need to delete
            if action == 'Move':
               treeCtrl.delete_Node(sourceData.nodeList, 'SearchClipNode')
               # Clear the Clipboard to prevent further Paste attempts, which are no longer valid as the SourceNode no longer exists!
               ClearClipboard()
            # Select the Destination Collection as the tree's Selected Item
            treeCtrl.SelectItem(destNode)

   # Drop a SearchClip on a SearchClip (Copy or Move a SearchClip into a position in the SortOrder)
   elif (sourceData.nodetype == 'SearchClipNode' and destNodeData.nodetype == 'SearchClipNode'):
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
         dlg = wx.MessageDialog(treeCtrl, prompt, _("Transana Confirmation"), style=wx.YES_NO | wx.ICON_QUESTION)
         # Display the Confirmation Dialog Box
         result = dlg.ShowModal()
         # Clean up the Confirmation Dialog Box
         dlg.Destroy()
         # If the user confirmed ...
         if result == wx.ID_YES:
            # ... we need to check to see if this Clip is a Duplicate, giving the user a chance to resolve the problem.
            (dupResult, newClipName) = CheckForDuplicateClipName(sourceData.text, treeCtrl, treeCtrl.GetItemParent(destNode))
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
                 prompt = unicode(_('%s cancelled for Clip "%s".  Duplicate Clip Name Error.'), 'utf8') % (copyMovePrompt, sourceData.text)
            else:
                 prompt = _('%s cancelled for Clip "%s".  Duplicate Clip Name Error.') % (copyMovePrompt, sourceData.text)
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
            # Now Add the new Node, using the SourceData's Data
            treeCtrl.add_Node('SearchClipNode', (_('Search'),) + nodeList, sourceData.recNum, sourceData.parent, False, insertPos=destNode)
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
    # Now put the data in the clipboard.
    wx.TheClipboard.SetData(cdo)

def CheckForDuplicateClipName(sourceClipName, treeCtrl, destCollectionNode):
   """ Check the destCollectionNode to see if sourceClipName already exists.  If so, prompt for a name change.
       Return True if duplicate is found, False if no duplicate is found or if Clip is renamed appropriately.  """
   # Before we do anything, let's make sure we have a Collection Node, not a Clip Node here
   if treeCtrl.GetPyData(destCollectionNode).nodetype == 'ClipNode':
       # If it's a clip, let's use its parent Collection
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
         
         # See if the current child is a Clip AND it has the same name as the source Clip
         if ((tempData.nodetype == 'ClipNode') or (tempData.nodetype == 'SearchClipNode')) and (treeCtrl.GetItemText(tempTreeItem).upper() == sourceClipName.upper()):
            # If so, prompt the user to change the Clip's Name.  First, build a Dialog to ask that question.
            dlg = wx.TextEntryDialog(TransanaGlobal.menuWindow, _('Duplicate Clip Name.  Please enter a new name for the Clip.'), _('Transana Error'), sourceClipName, style=wx.OK | wx.CANCEL | wx.CENTRE)
            # Position the Dialog Box in the center of the screen
            dlg.CentreOnScreen()
            # Show the Dialog Box
            dlgResult = dlg.ShowModal()
            # If the user selected OK AND changed the Clip Name ...
            if (dlgResult == wx.ID_OK) and (dlg.GetValue() != sourceClipName):
               # Let's look at the new name ...
               sourceClipName = dlg.GetValue()
               # ... and see if it is a Duplicate Clip Name by recursively calling this method
               (result, sourceClipName) = CheckForDuplicateClipName(sourceClipName, treeCtrl, destCollectionNode)
               # Clean up after prompting for the new Clip name
               dlg.Destroy()
               # We can stop looking for duplicate names.  We've already found it.
               break
            else:
               # If the user selected CANCEL or failed to change the Clip Name,
               # we set result to True to indicate that a Duplicate Clip Name was found.
               result = True
               # Clean up after prompting for a new Clip Name
               dlg.Destroy()
               # We can stop looking for duplicate names.  We've already found it.
               break

         # If we're not at the last child, get the next child for dropNode
         if tempTreeItem != treeCtrl.GetLastChild(destCollectionNode):
            (tempTreeItem, cookieVal) = treeCtrl.GetNextChild(destCollectionNode, cookieVal)
         # If we are at the last child, exit the while loop.
         else:
            break

   # Return the result and the Clip Name, as it could have been changed.
   return (result, sourceClipName)

def CopyMoveClip(treeCtrl, destNode, sourceClip, sourceCollection, destCollection, action):
   """ This function copies or moves sourceClip to destCollection, depending on the value of 'action' """
   if action == 'Copy':
      # Make a duplicate of the clip to be copied
      newClip = sourceClip.duplicate()
      # To place the copy in the destination collection, alter its Collection Number, Collection ID, and Sort Order value
      newClip.collection_num = destCollection.number
      newClip.collection_id = destCollection.id
   elif action == 'Move':
      # Lock the Clip Record to prevent other users from altering it simultaneously
      sourceClip.lock_record()
      # To move a clip, alter its Collection Number, Collection ID, and Sort Order value
      sourceClip.collection_num = destCollection.number
      sourceClip.collection_id = destCollection.id

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
   (dupResult, newClipName) = CheckForDuplicateClipName(clipName, treeCtrl, destNode)
   
   # If a Duplicate Clip Name is found and the error situation not resolved, show an Error Message
   if dupResult:
      # Unlock the source clip (before presenting the dialog to keep if from being locked by a slow user response.)
      if action == 'Move':
          sourceClip.unlock_record()
      # Report the failure to the user, although it's already known to have failed because they pressed "cancel".
      if 'unicode' in wx.PlatformInfo:
          # Encode with UTF-8 rather than TransanaGlobal.encoding because this is a prompt, not DB Data.
          prompt = unicode(_('%s cancelled for Clip "%s".  Duplicate Clip Name Error.'), 'utf8') % (action, sourceClip.id)
      else:
          prompt = _('%s cancelled for Clip "%s".  Duplicate Clip Name Error.') % (action, sourceClip.id)
      dlg = Dialogs.ErrorDialog(treeCtrl, prompt)
      dlg.ShowModal()
      dlg.Destroy()
      return None
   else:
      # The user may have given the clip a new Clip Name in CheckForDuplicateClipName.  
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
                 if DEBUG:
                     print 'Message to send = "RN %s"' % msg
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
         treeCtrl.add_Node('ClipNode', nodeList, newClip.number, newClip.collection_num)

         # Now let's communicate with other Transana instances if we're in Multi-user mode
         if not TransanaConstants.singleUserVersion:
            msg = "ACl %s"
            data = (nodeList[1],)

            for nd in nodeList[2:]:
               msg += " >|< %s"
               data += (nd, )

            if DEBUG:
               print 'DragAndDropObjects.CopyMoveClip(Copy): Message to send =', msg % data
                
            if TransanaGlobal.chatWindow != None:
               TransanaGlobal.chatWindow.SendMessage(msg % data)

         # When copying a Clip and setting its sort order, we need to keep working with the new clip
         # rather than the old one.  Having CopyClip return the new clip makes this easy.
         return newClip
      elif action == 'Move':
         nodeList = (_('Collections'), ) + nodeList + (sourceClip.id, )
         # Add the Node to the Tree
         treeCtrl.add_Node('ClipNode', nodeList, sourceClip.number, sourceClip.collection_num)
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

            if DEBUG:
               print 'DragAndDropObjects.CopyMoveClip(Move): Message to send =', msg % data
                
            if TransanaGlobal.chatWindow != None:
               TransanaGlobal.chatWindow.SendMessage(msg % data)

         # When copying a Clip and setting its sort order, we need to keep working with the new clip
         # rather than the old one.  Having CopyClip return the new clip makes this easy.
         return sourceClip

def ChangeClipOrder(treeCtrl, destNode, sourceClip, sourceCollection):
   """ This function changes the order of the clips in a Collection """

   # TODO:  Obtain Record Locks on all Clips here instead of below, and stop if they are not obtained.

   # If we are changing Clip Sort Order, the clip's Notes need to travel with the Clip.  The first step is to
   # get a list of those Notes.
   noteList = DBInterface.list_of_notes(Clip=sourceClip.number)

   # Remove the old Clip from the Tree   
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
   # Now add the Collections Root to the front of the Node List and the Clip to the end of the Node List   
   nodeList = (_('Collections'), ) + nodeList + (sourceClip.id, )   
   # Now request that the defined node be deleted   
   treeCtrl.delete_Node(nodeList, 'ClipNode')

   # Insert the new node in the proper position.  This is done by manipulating the tree directly, as "add_Node"
   # doesn't know about Sort Order
   # First, let's identify the Parent Node we're working with
   parentNode = treeCtrl.GetItemParent(destNode)

   # We need to figure out the position amongst the parentNode's children where we should drop the new node.
   # Initialize a variable to track this.
   nodeCounter = 0

   # What we need to do here is to iterate through all the Clips in a Collection, reassigning SortOrders
   # as we go, and insert the new record (both into the tree and into the SortOrder) as we go.  Since we
   # know the tree structure has the correct order and can tell us where to insert the new value, we will
   # use the Tree Structure for our iteration rather than going out to the DBInterface.
      
   # wxTreeCtrl requires the "cookie" value to list children.  Initialize it.
   cookie = 0
   # Get the first child of the dropNode 
   (tempNode, cookie) = treeCtrl.GetFirstChild(parentNode)

   # The node is getting inserted twice on the Mac.  Let's make sure that doesn't happen by tracking it explicitly.
   nodeInserted = False
      
   # Iterate through all the dropNode's children
   while tempNode.IsOk():
      # Get the current child's Node Data
      tempNodeData = treeCtrl.GetPyData(tempNode)
      # If we are looking at a Clip, and this is the Clip where the dropped item should be inserted ...
      if (not nodeInserted) and (tempNodeData.nodetype == 'ClipNode') and (treeCtrl.GetItemText(tempNode) == treeCtrl.GetItemText(destNode)):

         # Create a new node in the tree BEFORE the active node (as signalled by the nodeCounter counter), naming it after   
         # the Source Clip
         newNode = treeCtrl.InsertItemBefore(parentNode, nodeCounter, sourceClip.id)   
         # Identify this as a Clip node by assigning the appropriate NodeData   
         nodedata = DatabaseTreeTab._NodeData(nodetype='ClipNode', recNum=sourceClip.number, parent=sourceClip.collection_num)   
         # Associate the NodeData with the node   
         treeCtrl.SetPyData(newNode, nodedata)   
         # We know the new node is a Clip, so assign the proper image   
         treeCtrl.set_image(newNode, "Clip16")   
         
         # The node has now been inserted, and should not be inserted again.
         nodeInserted = True
         
         # If there are Clip Notes, we need to make sure they travel with the Clip
         if noteList != []:
             # We accomplish this using the TreeCtrl's "add_note_nodes" method
             treeCtrl.add_note_nodes(noteList, newNode, Clip=sourceClip.number)

         # Since we just inserted a new Node, we need to increment our NodeCounter
         nodeCounter += 1
         # Open a copy of the Clip that is being inserted
         localClip = Clip.Clip(sourceClip.number)
         # Lock the Clip Record
         localClip.lock_record()
         # Set the Clip's Sort Order based on the NodeCounter
         localClip.sort_order = nodeCounter
         # Save the Clip Record
         localClip.db_save()
         # Unlock the Clip Record
         localClip.unlock_record()

      # Increment the Node Counter
      nodeCounter += 1

      # If the current node is a Clip, let's reset its sort order
      if (tempNodeData.nodetype == 'ClipNode'):
         # Load the Clip
         localClip = Clip.Clip(tempNodeData.recNum)
         # Lock the Clip Record
         localClip.lock_record()
         # Reset the Sort Order based on the NodeCounter
         localClip.sort_order = nodeCounter
         # Save the Clip
         localClip.db_save()
         # Unlock the Clip Record
         localClip.unlock_record()

      # If we are looking at the last Child in the Parent's Node, exit the while loop
      if tempNode == treeCtrl.GetLastChild(parentNode):
          # We need to message the re-introduction of the node.
          # Now let's communicate with other Transana instances if we're in Multi-user mode
          if not TransanaConstants.singleUserVersion:
              msg = "AClSO %s"
              # We have a nodeList from deleting the node above.  We'll use that as the basis for the current nodeList,
              # but we need to insert the DropNode into it at the second-to-last position
              nodeList = nodeList[:-1] + (treeCtrl.GetItemText(destNode),) + (nodeList[-1],)
              data = (nodeList[1],)

              for nd in nodeList[2:]:
                  msg += " >|< %s"
                  data += (nd, )

              if DEBUG:
                  print 'DragAndDropObjects.CreateClip(): Message to send =', msg % data
                    
              if TransanaGlobal.chatWindow != None:
                  TransanaGlobal.chatWindow.SendMessage(msg % data)

          break
      # If not, load the next Child record
      else:
         (tempNode, cookie) = treeCtrl.GetNextChild(parentNode, cookie)

# Copyright (C) 2004 - 2006 The Board of Regents of the University of Wisconsin System 
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

"""This module reads a WAV file and creates a Waveform Graphics File """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "WaveformGraphic DEBUG is ON."

import wx          # Import wxWindows
import Dialogs
import wave                      # Import the wave module for processing Wave files
import sys                       # Import Python's sys module

def WaveformGraphicCreate(waveFilename, waveformFilename, startPoint, mediaLength, graphicSize, saveFile=False):
    try:

        if DEBUG:
            print "WaveformGraphicCreate() filename = '%s'" % waveFilename.encode('utf8')
            
        # Open the Wave File  (NOTE:  Python 2.4.2 can take a Unicode filename here.  Python 2.3.5 can't.)
        waveFile = wave.open(waveFilename, 'r')            
          
        if DEBUG:
            # Print the information from the Wave Header
            print '\n\nStart = ', startPoint, ', length = ', mediaLength
            print "Sampling Rate = %s" % waveFile.getframerate()
            print "Bytes per Sample = %s" % waveFile.getsampwidth()
            print "Channels = %s" % waveFile.getnchannels()
            print "total Frames = %s" % waveFile.getnframes()
            print "Compression Type = %s (%s)\n" % (waveFile.getcomptype(), waveFile.getcompname())

        # Added for Batch Waveform Generation, when we don't know the media file length
        if mediaLength == 0:
            mediaLength = waveFile.getnframes() * 1000

        # Create an Empty Bitmap
        theBitmap = wx.EmptyBitmap(graphicSize[0], graphicSize[1])
        # Create a Device Context on which to actually draw
        dc = wx.BufferedDC(None, theBitmap)
        # Set the background color of the Device Context
        dc.SetBackground(wx.Brush(wx.WHITE))
        # Clear the Device Context
        dc.Clear()
        # Create the pen for making the drawing.  We want to draw in RED, 1 pixel width line, solid line.
        pen = wx.Pen(wx.RED, 1, wx.SOLID)
        # Set the Pen for the Device Context
        dc.SetPen(pen)

        # Begin drawing to the Device Context
        dc.BeginDrawing()

        
        # Read the appropriate number of frames to position properly in the wave file
        # Number of seconds into the file * Frame Rate
        #print "read to ",float(startPoint) / 1000.0 * waveFile.getframerate(),"frames"
        frames = waveFile.readframes(float(startPoint) / 1000.0 * waveFile.getframerate())

        if DEBUG:
            print "read ",float(mediaLength)/1000.0 * waveFile.getframerate()
            
        totalFramesToRead = float(mediaLength)/1000.0 * waveFile.getframerate()

        # Calculate the number of WAVE data chunks to be read per line displayed in the graphic,
        # This value must be at least 1.
        ChunkSize = max(int(round(totalFramesToRead / graphicSize[0])), 1)

        if totalFramesToRead / graphicSize[0] < 1:
            print "\n\nTODO:  Zoomed in so that Number of Lines is less than Graphic Width!!\n\n"

        if DEBUG:
            print 'Chunksize = ', ChunkSize, '  (', totalFramesToRead, '/', graphicSize[0], ')', waveFile.getnframes()
            totalFramesRead = 0
        
        # Draw the actual WaveForm
        # for each pixel position in the graphic's width ...
        for loop in range(0, graphicSize[0]-1):
            # Read the appropriate number of chunks from the wave file
            frames = waveFile.readframes(ChunkSize)

            if DEBUG and (loop % 10 == 0):
                totalFramesRead += ChunkSize * 10  # * 10 because we're only reporting every 10th read!
                print "loop:  %d  totalFramesRead = %d  totalFrames = %d" % (loop, totalFramesRead, waveFile.getnframes())

            # Don't break all of Transana if we couldn't extract the wave
            if len(frames) == 0:
                break

            # Determine the largest value in the data read (this produces the best-looking graph!)
            frame = max(frames)
            # Process the data differently based on the Bytes Per Sample value of the Wave File
            if waveFile.getsampwidth() == 1:
                #This is for the 8-bit Bytes per Sample setting
                # The byte value represents sound, with 128 being silence and deviation from it being louder.
                # Therefore, determine the distance the Byte Value (frame) differs from silence.
                if ord(frame) > 128:
                   amplitude = ord(frame) - 128
                else:
                   amplitude = 128 - ord(frame)
                # Adjust the raw amplitude (0 .. 255 range) for the size of the graphic canvas
                amplitude = round(amplitude * graphicSize[1] / 256.0)
                # Determine the coordinates for drawing the amplitude line on the Device Context
                # The horizontal value equals the value of the loop
                x = loop
                # The vertical values represent the divergence of amplitude from the center of the graphic
                y1 = round((graphicSize[1]/2.0 - amplitude))
                y2 = round((graphicSize[1]/2.0 + amplitude))
                # draw the line on the Device Context
                dc.DrawLine(int(x), int(y1), int(x), int(y2))
            else:
                #This is for the 16-bit Bytes per Sample setting
                print "Waveform for 16-bit wave files not yet implemented."

        # Draw a black line down the center of the Waveform to show the true center
        # Create the pen for making the drawing.  We want to draw in RED, 1 pixel width line, solid line.
        pen = wx.Pen(wx.BLACK, 1, wx.SOLID)
        # Set the Pen for the Device Context
        dc.SetPen(pen)
        # Draw the line
        dc.DrawLine(0, int(round(graphicSize[1]/2.0)), int(graphicSize[0]-1), int(round(graphicSize[1]/2.0)))

        # Close the Wave File   
        waveFile.close()

        # Signal that drawing is complete
        dc.EndDrawing()

        if saveFile and theBitmap.Ok():
            # Save the Bitmap as a BMP file
            # theBitmap.SaveFile(waveformFilename, wx.BITMAP_TYPE_BMP)
            
            # Okay, on Win2K, we can't save BMPs because of a bug in wxPython.  Therefore,
            # I'm changing all of this from BMP to PNG format.  This change requires code changes
            # here, in VisualizationWindow.py, and in GraphicsControlClass.py.  Search for "png"
            # to find all the changes.
            theBitmap.SaveFile(waveformFilename[:-3]+'png', wx.BITMAP_TYPE_PNG)

        return True

    except:
        if DEBUG:
            import traceback
            traceback.print_exc(file=sys.stdout)

        # We need to just pass the exception on up.  This routine is called from an IDLE event, and that event needs to know
        # to stop trying to create this file!!
        raise

# Copyright (C) 2004 - 2014 The Board of Regents of the University of Wisconsin System 
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

# Import wxPython
import wx
# Import Transana's Dialogs
import Dialogs
# Import Python's sys module
import sys
# Import Python's wave module for processing Wave files
import wave


# import numpy


def WaveformGraphicCreate(waveFilename, waveformFilename, startPoint, mediaLength, graphicSize, colors = (wx.CYAN, wx.GREEN, wx.BLUE, wx.RED), style='waveform'):
    try:
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

        # Define the waveform colors, selecting just the number needed from the list, right-justified
        waveformColors = colors[-len(waveFilename):]
        # Initialize the Color Index
        colorIndex = 0

#        print "START"
        
        # Iterate through the wave files to be processed, in reverse order
        for wavFileIndex in range(len(waveFilename) - 1, -1, -1):
            # If the waveFilename entry doesn't HAVE a "Show" value, or if it has one that is set to True,
            # then draw this waveform.  ie default to show, only don't show if SHOW exists and is False.
            if (not waveFilename[wavFileIndex].has_key('Show')) or waveFilename[wavFileIndex]['Show']:
                # Get the wave file name
                wavFile = waveFilename[wavFileIndex]
                # Create a pen in the color this waveform should appear in
                pen = wx.Pen(waveformColors[colorIndex], 1, wx.SOLID)
                # Set the pen in the device context
                dc.SetPen(pen)
                
                # Open the Wave File  (NOTE:  Python 2.4.2 can take a Unicode filename here.  Python 2.3.5 can't.)
                waveFile = wave.open(wavFile['filename'], 'r')            
                  
                # Added for Batch Waveform Generation, when we don't know the media file length
                if mediaLength <= 0:
                    # Calculate it from the length of the wave file
                    mediaLength = waveFile.getnframes() * 1000

                # Read the appropriate number of frames to position properly in the wave file
                # Number of seconds into the file * Frame Rate

                # If we are at the beginning of the virtual media file ...
                if startPoint == 0:
                    # ... the start point for THIS media file needs to be adjusted for its offset
                    sp = int((float(wavFile['offset']) / float(mediaLength)) * (graphicSize[0] - 1))
                    # ... and the end point for THIS media file needs to be determined based on offset and length
                    ep = int((float(wavFile['offset'] + wavFile['length']) / float(mediaLength)) * (graphicSize[0] - 1)) + 2
                # If we are NOT at the beginning of the virtual media file ...
                else:
                    # ... Adjust the offset for THIS media file by the value of the Clip starting point

                    # Hmmmm.  I don't understand this.  If the offset is negative, we need to ignore it, as it shifts the
                    # waveform, but if it's positive, we need to compensate for it.

#                    if wavFile['offset'] < 0:
#                        print "***********     ALERT     WaveformGraphic.OnIdle() change     ALERT     ****************"
                        
#                    indent = max(0, wavFile['offset']) - startPoint
                    indent = wavFile['offset'] - startPoint

                    # If we have a positive value ...
                    if indent >= 0:
                        # ... then we can use that.
                        sp = int((float(indent) / float(mediaLength)) * (graphicSize[0] - 1))
                    # If we have a negative value (clip starts before this media file's start) ...
                    else:
                        # ... then set the media to the beginning.  It'll join in later.
                        sp = 0

                        if DEBUG:
                            print "read to ",float(abs(indent)) / 1000.0 * waveFile.getframerate(),"frames"
                            
                        # Indent the wave file the appropriate number of frames to get to the right part of the wave file
                        frames = waveFile.readframes(float(abs(indent)) / 1000.0 * waveFile.getframerate())

#                        print "**", startPoint, indent, float(abs(indent)) / 1000.0 * waveFile.getframerate(), float(indent) / 1000.0 * waveFile.getframerate()

                    # If we're in a clip, the ending point can be determined by looking at the waveform's WIDTH!!
                    ep = graphicSize[0] - 1

                # Calculate the total number of frames in the wave file
                totalFramesToRead = float(mediaLength)/1000.0 * waveFile.getframerate()

                # Calculate the number of WAVE data chunks to be read per line displayed in the graphic,
                # This value must be at least 1.
                ChunkSize = max(int(round(totalFramesToRead / graphicSize[0])), 1)

                if DEBUG and (totalFramesToRead / graphicSize[0] < 1):
                    print "\n\nTODO:  Zoomed in so that Number of Lines is less than Graphic Width!!\n\n"


                max1 = min1 = 0
                
                # Draw the actual WaveForm
                # for each pixel position in the graphic's width ...
                for loop in range(sp, ep):
                    # Read the appropriate number of chunks from the wave file
                    frames = waveFile.readframes(ChunkSize)

                    # Don't break all of Transana if we couldn't extract the wave
                    if len(frames) == 0:
                        break

                    # Determine the largest value in the data read (this produces the best-looking graph!)
                    frame = max(frames)
                    # Process the data differently based on the Bytes Per Sample value of the Wave File
                    if waveFile.getsampwidth() == 1:
                        if style == 'waveform':
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
                        elif style == 'spectrogram':

#                            print "Waveform style = spectrogram", sp, ep, ep-sp

#                            print frames, type(frames), len(frames)
#                            print

                            sigList = []
                            for loop2 in range(len(frames)):

                                val = ord(frames[loop2])
                                if val > 128:
                                    sigList.append(val - 128)
                                else:
                                    sigList.append(128 - val)

#                            print sigList
#                            print

                            max1 = max(max1, max(sigList))
                            min1 = min(min1, min(sigList))

                            sig = numpy.array(sigList)
                            
#                            print sig
#                            print

                            spectrum = 10*numpy.log10(abs(numpy.fft.rfft(sig)))

#                            print spectrum, len(spectrum)

#                            print "Max =", max1, "Min =", min1

                            x = loop
                            for loop2 in range(len(spectrum)):

#                                print loop2, spectrum[loop2], type(spectrum[loop2]),

                                if spectrum[loop2] in [numpy.inf, -numpy.inf]:
                                    n = 0
                                else:
                                    n = max(0, int(5 * spectrum[loop2]))

#                                    print 5 * spectrum[loop2], n

                                    n = max(0, int(5 * spectrum[loop2]))
                                
                                pen.SetColour(wx.Colour(255-n, 255-n, 255-n))
                                dc.SetPen(pen)
                                dc.DrawPoint(x, loop2)
                                              
                        
                    else:
                        #This is for the 16-bit Bytes per Sample setting
                        print "Waveform for 16-bit wave files not yet implemented."

                # Close the Wave File   
                waveFile.close()
            # Iterate the color index, so the next waveform will be in the next color
            colorIndex += 1

        # Draw a black line down the center of the Waveform to show the true center
        # Create the pen for making the drawing.  We want to draw in RED, 1 pixel width line, solid line.
        pen = wx.Pen(wx.BLACK, 1, wx.SOLID)
        # Set the Pen for the Device Context
        dc.SetPen(pen)
        # Draw the line
        dc.DrawLine(0, int(round(graphicSize[1]/2.0)), int(graphicSize[0]-1), int(round(graphicSize[1]/2.0)))

        # Signal that drawing is complete
        dc.EndDrawing()


#        print "END"
        
        # If the waveformFilename is ":memory:", we return an Image
        if waveformFilename == ':memory:':
            # If we have a good Bitmap ...
            if theBitmap.Ok():
                # ... convert it to an Image and return it
                return theBitmap.ConvertToImage()
            # If we do NOT have a good image ...
            else:
                # ... return None to signal failure
                return None
        # If the waveformFilename is NOT ":memory:", we save the bitmap as a PNG file
        else:
            # If we have a good Bitmap ...
            if theBitmap.Ok():
                # ... save the Bitmap as a PNG file ...
                theBitmap.SaveFile(waveformFilename[:-3]+'png', wx.BITMAP_TYPE_PNG)
                # ... and return True to signal success
                return True
            # If we do NOT have a good image ...
            else:
                # ... return False to signal failure
                return False

    except:
        if DEBUG:
            import traceback
            traceback.print_exc(file=sys.stdout)

        # We need to just pass the exception on up.  This routine is called from an IDLE event, and that event needs to know
        # to stop trying to create this file!!
        raise

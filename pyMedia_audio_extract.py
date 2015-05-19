# Copyright (C) 2005 - 2006  The Board of Regents of the University of Wisconsin System 
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

""" This module implements waveform extraction using pyMedia """

__author__ = 'David K. Woods <dwoods@wcer.wisc.edu>'

DEBUG = False
if DEBUG:
    print "pyMedia_audio_extract DEBUG is ON."

import os
import sys
import time
import wave
import pymedia

def ExtractWaveFile(filename, outputFilename, feedbackDialog=None):

    if os.path.exists(filename):
        if DEBUG:
            print "File exists"
    else:
        if DEBUG:
            print "File does not exist"
        return None

    fileext = filename.split('.')[-1].lower()    # fileroot, fileext = os.path.splitext(filename)

    if DEBUG:
        print "File Extension = ", fileext

    audioExtensions = ['mp3', 'wav', 'wma']
    videoExtensions = ['avi', 'mpeg', 'mpg']
    registeredExtensions = audioExtensions + videoExtensions

    if fileext in videoExtensions:

        if DEBUG:
            print 'file "%s" is a Video File' % filename.encode('utf8')

        outputFilename = outputFilename + '.tmp.wav'

        # Define a demuxer object
        demuxer = pymedia.muxer.Demuxer( fileext )

        # Open the Video File
        videoFile = open(filename, 'rb')

        fileSize = os.path.getsize(filename)
        
        if DEBUG:
            print "file size =", fileSize

        outputFile = None

        # Read the first sample from the file
        sample = videoFile.read( 400000 )

        bytesRead = len(sample)

        # Use the demuxer to parse the first sample into raw data streams
        rawData = demuxer.parse(sample)

        audioStreams = None
        for s in demuxer.streams:
            if s and s['type'] == pymedia.muxer.CODEC_TYPE_AUDIO:
                audioStreams = s
                break

        if audioStreams == None:

            if DEBUG:
                print "No Audio Streams"

            sys.exit(1)

        # Use the index from the first audio stream to determine the stream ID for the audio stream
        audio_id = audioStreams['index']

        if DEBUG:
            print "Audio Stream at index %d" % audio_id

        audio_params = demuxer.streams[audio_id]

        # Set an Audio Decoder to the appropriate stream
        if demuxer.streams[audio_id]['id'] != 49:

            if DEBUG:
                print "id != 49"
            
            decoder = pymedia.audio.acodec.Decoder( demuxer.streams[audio_id] )

            # while there's data to be processed ...
            while len(sample) > 0:
                for frames in rawData:
                    if frames[0] == audio_id:
                        data = decoder.decode(frames[1])

                        if data:

                            if DEBUG:
                                print "audio_params =", audio_params
                                print

                                print "channels:", data.channels
                                print "sampwidth:", (data.bitrate / data.sample_rate) / 8  # NO NO NO
                                print "bitrate:", data.bitrate
                                print "framerate:", data.sample_rate
                                
                            if outputFile == None:
                                # Open the Wave File
                                # NOTE:  Unicode requires that the opening of the file be seperate from the wave.open call.
                                #        wave.open requires either a str variable (not unicode) or a file object as the first
                                #        parameter.
                                waveFileObject = file(outputFilename, 'wb')
                                outputFile = wave.open(waveFileObject, 'wb')
                                outputFile.setparams((data.channels, 2,
                                                      data.sample_rate, 0, 'NONE', ''))

                        if data:
                            outputFile.writeframes(data.data)
                    
                if DEBUG:
                    print "%12d of %12d (%5.2f)" % (bytesRead, fileSize, ((100.0 * bytesRead)/fileSize))

                if feedbackDialog != None:
                    feedbackDialog.Update(int((50.0 * bytesRead)/fileSize), 0)

                sample = videoFile.read( 400000 )
                bytesRead += len(sample)
                rawData = demuxer.parse(sample)

        else:

            if DEBUG:
                print "id == 49"
            
                print demuxer.streams[audio_id]

            if outputFile == None:
                # Open the Wave File
                # NOTE:  Unicode requires that the opening of the file be seperate from the wave.open call.
                #        wave.open requires either a str variable (not unicode) or a file object as the first
                #        parameter.
                waveFileObject = file(outputFilename, 'wb')
                outputFile = wave.open(waveFileObject, 'wb')
                outputFile.setparams((demuxer.streams[audio_id]['channels'],
                                      demuxer.streams[audio_id]['bitrate'] / demuxer.streams[audio_id]['sample_rate'] / 8,
                                      demuxer.streams[audio_id]['sample_rate'], 0, 'NONE', ''))

            # while there's data to be processed ...
            while len(sample) > 0:
                for frames in rawData:
                    if frames[0] == 1:
                        outputFile.writeframes(frames[1])

                if DEBUG:
                    print "%12d of %12d (%5.2f)" % (bytesRead, fileSize, ((100.0 * bytesRead)/fileSize))

                if feedbackDialog != None:
                    feedbackDialog.Update(int((50.0 * bytesRead)/fileSize), 0)

                sample = videoFile.read( 400000 )
                bytesRead += len(sample)
                rawData = demuxer.parse(sample)
                


        videoFile.close()

        outputFile.close()

        return outputFilename

    elif (fileext in audioExtensions) and (fileext != 'wav'):

        if DEBUG:
            print 'file "%s" is a non-WAV Audio File' % filename.encode('utf8')

        outputFilename = outputFilename + '.tmp.wav'

        # Open the Audio File
        audioFile = open(filename, 'rb')

        fileSize = os.path.getsize(filename)
        if DEBUG:
            print "file size =", fileSize

        outputFile = None

        # Create a demuxer to process the file
        demuxer = pymedia.muxer.Demuxer( fileext )

        # Read the first sample from the file
        sample = audioFile.read( 400000 )

        bytesRead = len(sample)

        # Initialize Decoder, used to signal that a real decode hasn't been created yet.
        decoder = None

        # while there's data to be processed ...
        while len(sample) > 0:

            frames = demuxer.parse( sample )
            for frame in frames:
                if decoder == None:
                    # Open a Decoder
                    decoder = pymedia.audio.acodec.Decoder( demuxer.streams[ 0 ] )
                    
                data = decoder.decode( frame[1] )

                if outputFile == None:
                    # Open the Wave File
                    # NOTE:  Unicode requires that the opening of the file be seperate from the wave.open call.
                    #        wave.open requires either a str variable (not unicode) or a file object as the first
                    #        parameter.
                    waveFileObject = file(outputFilename, 'wb')
                    outputFile = wave.open(waveFileObject, 'wb')

                    if DEBUG:
                        print "data:"
                        print "  sample_rate:", data.sample_rate
                        print "  bitrate:", data.bitrate
                        print "  channels:", data.channels

                    outputFile.setparams((data.channels, 2, data.sample_rate, 0, 'NONE', ''))

                    
                if data and data.data:
                    outputFile.writeframes(data.data)

            if DEBUG:
                print "%12d of %12d (%5.2f)" % (bytesRead, fileSize, ((100.0 * bytesRead)/fileSize))

            if feedbackDialog != None:
                feedbackDialog.Update(int((50.0 * bytesRead)/fileSize), 0)

            sample = audioFile.read( 400000 )
            bytesRead += len(sample)


        audioFile.close()

        outputFile.close()

        return outputFilename

    elif fileext == 'wav':

        if DEBUG:
            print "File is a WAVE file.  No need to do anything!"

        if feedbackDialog != None:
            feedbackDialog.Update(50, 0)

        return filename


def DecimateWaveFile(waveFilename, outputFilename, decimationLevel, feedbackDialog=None):

    # Open the Audio File
    # Open the Wave File
    # NOTE:  Unicode requires that the opening of the file be seperate from the wave.open call.
    #        wave.open requires either a str variable (not unicode) or a file object as the first
    #        parameter.
    waveFileObject = file(waveFilename, 'rb')
    audioFile = wave.open(waveFileObject, 'rb')

    fileSize = os.path.getsize(waveFilename)
    totalFrames = audioFile.getnframes()

    if DEBUG:
        print "file size =", fileSize, totalFrames

    # Open the Wave File
    # NOTE:  Unicode requires that the opening of the file be seperate from the wave.open call.
    #        wave.open requires either a str variable (not unicode) or a file object as the first
    #        parameter.
    waveFileObject = file(outputFilename, 'wb')
    outputFile = wave.open(waveFileObject, 'wb')

    if DEBUG:
        print audioFile.getparams()

    import time

    startTime = time.time()

    nextDisplayIndex = 2
    cnt = 0

    if DEBUG:
        print "sample width =", audioFile.getsampwidth() * 8,
        if audioFile.getnchannels() == 1:
            print "mono"
        else:
            print "stereo"

        print "Total frames =", audioFile.getnframes()
    
    blocksize = 32 * audioFile.getsampwidth() * 8 * decimationLevel

    if DEBUG:
        print "block size =", blocksize

    framerate = audioFile.getframerate()

    if DEBUG:
        print "framerate =",  framerate, '  decimationLevel =', decimationLevel

    if (framerate < 44100) and (decimationLevel >= 2):
        decimationLevel /= 2
    if (framerate < 22050) and (decimationLevel >= 2):
        decimationLevel /= 2
    if (framerate < 11025) and (decimationLevel >= 2):
        decimationLevel /= 2

    if DEBUG:
        print 'new decimationLevel =', decimationLevel

#    outputFile.setparams((audioFile.getnchannels(), audioFile.getsampwidth(), audioFile.getframerate() / decimationLevel, 0, 'NONE', ''))
    outputFile.setparams((1, 1, audioFile.getframerate() / decimationLevel, 0, 'NONE', ''))

    for index in range(totalFrames / blocksize):

        frame = audioFile.readframes(blocksize)
        frameLength = len(frame) / blocksize

        for fr in range(0, len(frame), decimationLevel * frameLength):
            if cnt < 100:
                cnt += 1
            data = frame[fr : fr + frameLength]
            # We need to convert the file from whatever it is to 8-bit mono.
            #
            # If we're dealing with 16-bit sound...
            if audioFile.getsampwidth() == 2:
                # ... and stereo sound ...
                if audioFile.getnchannels() == 2:
                    # I have NO idea why this works.
                    v1 = ord(data[1])
                    v2 = ord(data[3])

                    # print val
                    
                    if v1 >= 128:
                        v1 = v1 - 128
                    else:
                        v1 = 128 + v1
                    if v2 >= 128:
                        v2 = v2 - 128
                    else:
                        v2 = 128 + v2
                    val = (v1 + v2) / 2
                    data = chr(val)
                # ... and mono sound ...
                else:
                    # I have NO idea why this works.
                    val = ord(data[1])
                    if val >= 128:
                        val = val - 128
                    else:
                        val = 128 + val
                    data = chr(val)
            # Else if we're dealing with 8-bit already ...
            else:
                # ... but the file is stereo ...
                if audioFile.getnchannels() == 2:
                    # ... we'll average the two samples
                    data = chr((ord(data[0]) + ord(data[1])) / 2)
                    
            # this if reduces 16-bit files to 8 bits
            if fr % audioFile.getsampwidth() == 0:
                outputFile.writeframes(data)

        if int((index * blocksize * 100.0) / totalFrames) >= nextDisplayIndex:
            if DEBUG:
                print index * blocksize, int((index * blocksize * 100.0)/ totalFrames), nextDisplayIndex

            if feedbackDialog != None:
                feedbackDialog.Update(int((nextDisplayIndex / 2) + 50), 0)
            nextDisplayIndex += 2



    if DEBUG:
        print "read =", time.time() - startTime

    audioFile.close()

    outputFile.close()

if __name__ == '__main__':
    # Works with Win MPEG-1
    # Works with Win MPEG-2
    # Works with Win and Mac AVI
    # Works with MP3
    # Works with WAVE

    # filename = "C:\Data\Video\Demo\Demo.mpg"
    # filename = "C:\Data\Video\DrugsINeed-MPEG1.mpg"
    # filename = "C:\Data\Video\DrugsINeed-MPEG2.mpg"
    # filename = "C:\\Data\\Video\\DrugsINeed.avi"
    # filename = "C:\\My Music\\A Fine Mess\\Bag O' Mess\\01Systematic.MP3"
    # filename = "C:\\My Music\\Unknown Artist\\David's Dream CD\\06 Dream CD 6 - Car Wheels on a Grav.wav"
    # filename = "C:\\My Music\\Unknown Artist\\David's Dream CD\\09 Dream CD 9 - DJ Talk.wav"
    # filename = "C:\Documents and Settings\DavidWoods\Desktop\WS_10004.WMA"
    # filename = "C:\\Documents and Settings\\DavidWoods\\Desktop\\testWave.wav"
    # filename = "C:\\Documents and Settings\\DavidWoods\\Desktop\\Benedictus.mp3"
    # filename = "C:\\Documents and Settings\\DavidWoods\\Desktop\\Benedictus.m4a"  NOT SUPPORTED
    # filename = "C:\\Documents and Settings\\DavidWoods\\Desktop\\sample44k8bitMo.wav"
    # filename = "C:\\Documents and Settings\\DavidWoods\\Desktop\\sample44k8bitSt.wav"
    # filename = "C:\\Documents and Settings\\DavidWoods\\Desktop\\sample44k16bitMo.wav"
    # filename = "C:\\Documents and Settings\\DavidWoods\\Desktop\\sample44k16bitSt.wav"

    filename = "e:\\Video\\Music\\03 Stacy's Mom.MP3"
    # filename = "e:\\Video\\Music\\4. Time Of Our Lives.mp3"
    # filename = "e:\\Video\\Demo\\Demo.mpg"

    if os.path.exists(filename):
        print "File exists"
    else:
        print "File does not exist"

    output = 'C:\\Documents and Settings\\DavidWoods\\Desktop\\Test.wav'

    tempFile = ExtractWaveFile(filename, output)
    if tempFile != None:
        DecimateWaveFile(tempFile, output, 16)


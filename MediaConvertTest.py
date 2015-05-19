# -*- coding: cp1252 -*-

__author__ = 'David Woods <dwoods@wcer.wisc.edu>'

# import wxPython
import wx

# Import Python's os and sys modules
import os, sys

class MediaConvert(wx.Dialog):
    """ Transana's Media Conversion Tool Dialog Box. """
    
    def __init__(self, parent, fileName=u'Demö.mpg'):
        """ Initialize the MediaConvert Dialog Box. """

        wx.SetDefaultPyEncoding('utf_8')

        self.locale = wx.Locale(parent.locale.Language)

        # Remember the File Name passed in
        self.fileName = fileName
        # Initialize the process variable
        self.process = None

        # Create the Dialog
        wx.Dialog.__init__(self, parent, -1, _('Media File Conversion'), size=wx.Size(600, 700), style=wx.DEFAULT_DIALOG_STYLE | wx.THICK_FRAME)

        # To look right, the Mac needs the Small Window Variant.
        if "__WXMAC__" in wx.PlatformInfo:
            self.SetWindowVariant(wx.WINDOW_VARIANT_SMALL)

        # Create the main Sizer, which will hold the box1, box2, etc. and boxButton sizers
        box = wx.BoxSizer(wx.VERTICAL)

        # Add the Source File label
        lblSource = wx.StaticText(self, -1, _("Source Media File:"))
        box.Add(lblSource, 0, wx.TOP | wx.LEFT | wx.RIGHT, 10)
        
        # Create the box1 sizer, which will hold the source file and its browse button
        box1 = wx.BoxSizer(wx.HORIZONTAL)

        # Create the Source File text box
        self.txtSrcFileName = wx.TextCtrl(self, -1, fileName)
        self.txtSrcFileName.SetValue(self.fileName)

        box1.Add(self.txtSrcFileName, 1, wx.EXPAND)
        # Spacer
        box1.Add((4, 0))
        # Add the Source Sizer to the Main Sizer
        box.Add(box1, 0, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Add the Information label
        lblMemo = wx.StaticText(self, -1, _("Information:"))
        box.Add(lblMemo, 0, wx.LEFT | wx.RIGHT, 10)

        # Add the Information text control
        self.memo = wx.TextCtrl(self, -1, style = wx.TE_MULTILINE)
        box.Add(self.memo, 1, wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Create the boxButtons sizer, which will hold the dialog box's buttons
        boxButtons = wx.BoxSizer(wx.HORIZONTAL)

        # Create a Close button
        btnClose = wx.Button(self, wx.ID_CANCEL, _("Close"))
        boxButtons.Add(btnClose, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.RIGHT, 10)

        # Add the boxButtons sizer to the main box sizer
        box.Add(boxButtons, 0, wx.ALIGN_RIGHT | wx.ALIGN_BOTTOM | wx.LEFT | wx.RIGHT | wx.BOTTOM, 10)

        # Define box as the form's main sizer
        self.SetSizer(box)
        # Set this as the minimum size for the form.
        self.SetSizeHints(minW = self.GetSize()[0], minH = int(self.GetSize()[1] * 0.75))
        # Tell the form to maintain the layout and have it set the intitial Layout
        self.SetAutoLayout(True)
        self.Layout()
        # Position the form in the center of the screen
        self.CentreOnScreen()

        # If a file name was passed in as a parameter ...
        if (self.fileName != ''):
            # ... process that file to prepare for conversion
            self.ProcessMediaFile(self.fileName)

    def ProcessMediaFile(self, inputFile):
        """ Process a Media File to see what it's made of, a pre-requisite to converting it.
            This process also populates the form's options. """
        
        # If we're messing with wxProcess, we need to define a function to clean up if we shut down!
        def __del__(self):

            # If a process has been defined ...
            if self.process is not None:
                # ... detach it
                self.process.Detach()
                #  ... Close its output
                self.process.CloseOutput()
                # ... and de-reference it
                self.process = None

        # Reset (re-initialize) all media file variables
        self.process = None

        # Clear the Information box
        self.memo.Clear()

        # Be prepared to capture the wxProcess' EVT_END_PROCESS
        self.Bind(wx.EVT_END_PROCESS, self.OnEndProcess)
        
        # Windows requires that we change the default encoding for Python for the audio extraction code to work
        # properly with Unicode files (!!!)  This isn't needed on OS X, as its default file system encoding is utf-8.
        # See python's documentation for sys.getfilesystemencoding()
        if 'wxMSW' in wx.PlatformInfo:
            # Set the Python Encoding to match the File System Encoding
            wx.SetDefaultPyEncoding(sys.getfilesystemencoding())
        # Just use the File Name, no encoding needed
        tempMediaFilename = inputFile

        self.memo.AppendText("wx version:  %s\n" % wx.VERSION_STRING)

        self.memo.AppendText("Encoding Information:\n  defaultPyEncoding: %s, filesystemencoding: %s\n\n" % (wx.GetDefaultPyEncoding(), sys.getfilesystemencoding()))
        self.memo.AppendText("Locale: %s  %s  %s\n\n" % (self.locale.Locale, self.locale.Name, self.locale.Language))

        process = '"echo" "%s"'
            
        # Create a wxProcess object
        self.process = wx.Process(self)
        # Call the wxProcess Object's Redirect method.  This allows us to capture the process's output!
        self.process.Redirect()
        # Encode the filenames to UTF8 so that unicode files are handled properly
        process = process.encode('utf8')

        self.memo.AppendText("Media Filename:\n")
        self.memo.AppendText("%s (%s)\n" % (tempMediaFilename, type(tempMediaFilename)))
        self.memo.AppendText("Process call:\n")
        self.memo.AppendText("%s (%s - %s)\n\n" % (process % tempMediaFilename, type(process), type(process % tempMediaFilename)))

        line = tempMediaFilename.encode('utf8')
        for tmpX in range(len(line)):
            self.memo.AppendText("%02d\t%03d" % (tmpX, ord(line[tmpX])))
            if ord(line[tmpX]) < 255:
                self.memo.AppendText(u"\t'%s'" % (unichr(ord(line[tmpX]))))
            self.memo.AppendText("\n")

        try:
            self.memo.AppendText("%02d  %s  (%s)\n\n" % (len(line), line, type(line)))
        except:
            self.memo.AppendText("EXCEPTION RAISED:\n%s\n%s\n\n" % (sys.exc_info()[0], sys.exc_info()[1]))

        # Call the Audio Extraction program using wxExecute, capturing the output via wxProcess.  This call MUST be asynchronous. 
        self.pid = wx.Execute(process % tempMediaFilename.encode('utf8'), wx.EXEC_ASYNC, self.process)

        # On Windows, we need to reset the encoding to UTF-8
        if 'wxMSW' in wx.PlatformInfo:
            wx.SetDefaultPyEncoding('utf_8')

    def OnEndProcess(self, event):
        """ End of wxProcess Event Handler """
        # If a process is defined ...
        if self.process is not None:

            # Get the Process' Input Stream
            stream = self.process.GetInputStream()
            # If that stream can be read ...
            if stream.CanRead():

                stream.flush()

                # ... read it!
                text = stream.read()

                self.memo.AppendText("stream.CanRead() call successful. (%s)\n\n" % type(text))
                    
                # Divide the text up into separate lines
                text = text.replace('\r\n', '\n')
                text = text.split('\n')
                # Process the input stream text one line at a time
                for line in text:

                    for tmpX in range(len(line)):
                        self.memo.AppendText("%02d\t%03d" % (tmpX, ord(line[tmpX])))
                        if ord(line[tmpX]) < 255:
                            self.memo.AppendText(u"\t'%s'" % (unichr(ord(line[tmpX]))))
                        self.memo.AppendText("\n")

                    try:
                        self.memo.AppendText("%02d  %s  (%s)\n\n" % (len(line), line, type(line)))
                    except:
                        self.memo.AppendText("EXCEPTION RAISED:\n%s\n%s\n\n" % (sys.exc_info()[0], sys.exc_info()[1]))

            # Since the process has ended, destroy it.
            self.process.Destroy()
            # De-reference the process
            self.process = None

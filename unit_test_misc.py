import datetime
import wx

import gettext

# This module expects i18n.  Enable it here.
__builtins__._ = wx.GetTranslation


import Dialogs
import Misc
import TransanaConstants


MENU_FILE_EXIT = 101

class FormCheck(wx.Frame):
    """ This window displays a variety of GUI Widgets. """
    def __init__(self,parent,id,title):

        wx.SetDefaultPyEncoding('utf8')

        wx.Frame.__init__(self,parent,-1, title, size = (800,600), style=wx.DEFAULT_FRAME_STYLE|wx.NO_FULL_REPAINT_ON_RESIZE)
        self.testsRun = 0
        self.testsSuccessful = 0
        self.testsFailed = 0

        mainSizer = wx.BoxSizer(wx.VERTICAL)

        self.txtCtrl = wx.TextCtrl(self, -1, "Unit Test:  Miscellaneous Functions\n\n", style=wx.TE_LEFT | wx.TE_MULTILINE)
        self.txtCtrl.AppendText("Transana Version:  %s     singleUserVersion:  %s\n\n" % (TransanaConstants.versionNumber, TransanaConstants.singleUserVersion))

        mainSizer.Add(self.txtCtrl, 1, wx.EXPAND | wx.ALL, 5)

        self.SetSizer(mainSizer)
        self.SetAutoLayout(True)
        self.Layout()
        self.CenterOnScreen()
        # Status Bar
        self.CreateStatusBar()
        self.SetStatusText("")
        self.Show(True)
        self.RunTests()

    def RunTests(self):
        # Tests defined:
        testsNotToSkip = []
        startAtTest = 100  # Should start at 1, not 0!
        endAtTest = 500   # Should be one more than the last test to be run!
        testsToRun = testsNotToSkip + range(startAtTest, endAtTest)

        t = datetime.datetime.now()

        if 10 in testsToRun:
            # dt_to_datestr
            testName = 'Misc.dt_to_datestr()'
            self.SetStatusText(testName)
            prompt = 'Raw Date/Time information:  %s\n\n  Is the correct M/D/Y formatted date string "%s"?' % (t, Misc.dt_to_datestr(t))
            self.ConfirmTest(prompt, testName)

        if 20 in testsToRun:
            # time_in_ms_to_str
            testName = 'Misc.time_in_ms_to_str()'
            self.SetStatusText(testName)
            prompt = 'Are the following time to string conversions correct?\n\nTime in ms:\ttime_in_ms_to_str():\n'
            for x in [100, 200, 1000, 2000, 10000, 20000, 60000, 120000, 600000, 1200000, 3600000, 7200000]:
                prompt += "  %12d\t  %s\n" % (x, Misc.time_in_ms_to_str(x))
            self.ConfirmTest(prompt, testName)

        if 30 in testsToRun:
            # TimeMsToStr
            testName = 'Misc.TimeMsToStr()'
            self.SetStatusText(testName)
            prompt = 'Are the following time to string conversions correct?\n\nTime in ms:\tTimeMsToStr():\n'
            for x in [1000, 2000, 10000, 20000, 60000, 120000, 600000, 1200000, 3600000, 7200000]:
                prompt += "  %12d\t  %s\n" % (x, Misc.TimeMsToStr(x))
            self.ConfirmTest(prompt, testName)

        if 40 in testsToRun:
            # time_in_str_to_ms
            testName = 'Misc.time_in_str_to_ms()'
            self.SetStatusText(testName)
            prompt = 'Are the following string to time conversions correct?\n\nTime in string:\ttime_in_str_to_ms():\n'
            for x in ['0:00:00.1', '0:00:00.2', '0:00:01.0', '0:00:02.0', '0:00:10.0', '0:00:20.0', '0:01:00.0', '0:02:00.0', '0:10:00.0', '0:20:00.0', '1:00:00.0', '2:00:00.0']:
                prompt += "  %s\t  %12d\n" % (x, Misc.time_in_str_to_ms(x))
            self.ConfirmTest(prompt, testName)

        if 100 in testsToRun:
            # English
            testName = 'English translation'

            langName = 'en'
            lang = wx.LANGUAGE_ENGLISH
            self.presLan_en = gettext.translation('Transana', 'locale', languages=['en']) # English
            self.presLan_en.install()
            self.locale = wx.Locale(lang)
            self.locale.AddCatalog("Transana")
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 110 in testsToRun:
            # Arabic
            testName = 'Arabic translation'
            langName = 'ar'
            lang = wx.LANGUAGE_ARABIC
            self.presLan_ar = gettext.translation('Transana', 'locale', languages=['ar']) # Arabic
            self.presLan_ar.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)
            
        if 120 in testsToRun:
            # Danish
            testName = 'Danish translation'
            langName = 'da'
            lang = wx.LANGUAGE_DANISH
            self.presLan_da = gettext.translation('Transana', 'locale', languages=['da']) # Danish
            self.presLan_da.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 130 in testsToRun:
            # German
            testName = 'German translation'
            langName = 'de'
            lang = wx.LANGUAGE_GERMAN
            self.presLan_de = gettext.translation('Transana', 'locale', languages=['de']) # German
            self.presLan_de.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 140 in testsToRun:
            # Spanish
            testName = 'Spanish translation'
            langName = 'es'
            lang = wx.LANGUAGE_SPANISH
            self.presLan_es = gettext.translation('Transana', 'locale', languages=['es']) # Spanish
            self.presLan_es.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 150 in testsToRun:
            # French
            testName = 'French translation'
            langName = 'fr'
            lang = wx.LANGUAGE_FRENCH
            self.presLan_fr = gettext.translation('Transana', 'locale', languages=['fr']) # French
            self.presLan_fr.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 160 in testsToRun:
            # Italian
            testName = 'Italian translation'
            langName = 'it'
            lang = wx.LANGUAGE_ITALIAN
            self.presLan_it = gettext.translation('Transana', 'locale', languages=['it']) # Italian
            self.presLan_it.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 170 in testsToRun:
            # Dutch
            testName = 'Dutch translation'
            langName = 'nl'
            lang = wx.LANGUAGE_DUTCH
            self.presLan_nl = gettext.translation('Transana', 'locale', languages=['nl']) # Dutch
            self.presLan_nl.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 180 in testsToRun:
            # Norwegian Bokmal
            testName = 'Norwegian Bokmal translation'
            langName = 'nb'
            if 'wxMac' in wx.PlatformInfo:
                lang = wx.LANGUAGE_ENGLISH
            else:
                lang = wx.LANGUAGE_NORWEGIAN_BOKMAL
            self.presLan_nb = gettext.translation('Transana', 'locale', languages=['nb']) # Norwegian Bokmal
            self.presLan_nb.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 190 in testsToRun:
            # Norwegian Nynorsk
            testName = 'Norwegian Nynorsk translation'
            langName = 'nn'
            if 'wxMac' in wx.PlatformInfo:
                lang = wx.LANGUAGE_ENGLISH
            else:
                lang = wx.LANGUAGE_NORWEGIAN_NYNORSK
            self.presLan_nn = gettext.translation('Transana', 'locale', languages=['nn']) # Norwegian Nynorsk
            self.presLan_nn.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 200 in testsToRun:
            # Swedish
            testName = 'Swedish translation'
            langName = 'sv'
            lang = wx.LANGUAGE_SWEDISH
            self.presLan_sv = gettext.translation('Transana', 'locale', languages=['sv']) # Swedish
            self.presLan_sv.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)

        if 210 in testsToRun:
            # Chinese Simplified
            testName = 'Chinese - Simplified translation'
            langName = 'zh'
            lang = wx.LANGUAGE_CHINESE
            self.presLan_zh = gettext.translation('Transana', 'locale', languages=['zh']) # Chinese Simplified
            self.presLan_zh.install()
            self.locale = wx.Locale(lang)
            prompt  = '%s:  Series     = %s (%s)\n' % (langName, _('Series'),     type(_('Series')))
            prompt += '     Collection = %s (%s)\n' % (          _('Collection'), type(_('Collection')))
            prompt += '     Keyword    = %s (%s)\n' % (          _('Keyword'),    type(_('Keyword')))
            prompt += '     Search     = %s (%s)\n' % (          _('Search'),     type(_('Search')))
            self.ConfirmTest(prompt, testName)


    def ShowMessage(self, msg):
        dlg = Dialogs.InfoDialog(self, msg)
        dlg.ShowModal()
        dlg.Destroy()

    def ConfirmTest(self, msg, testName):
        dlg = Dialogs.QuestionDialog(self, msg, "Unit Test 3: Miscellaneous Functions", noDefault=True)
        result = dlg.LocalShowModal()
        dlg.Destroy()
        self.testsRun += 1
        self.txtCtrl.AppendText('Test "%s" ' % testName)
        if result == wx.ID_YES:
            self.txtCtrl.AppendText('Passed.')
            self.testsSuccessful += 1
        else:
            self.txtCtrl.AppendText('FAILED.')
            self.testsFailed += 1
        self.txtCtrl.AppendText('\nTotal Tests Run:  %d  Tests passes:  %d  Tests failed:  %d.\n\n' % (self.testsRun, self.testsSuccessful, self.testsFailed))




class MyApp(wx.App):
   def OnInit(self):
      frame = FormCheck(None, -1, "Unit Test 1: Objects and Forms")
      self.SetTopWindow(frame)
      return True
      

app = MyApp(0)
app.MainLoop()

import gettext
import os
import wx

__builtins__._ = wx.GetTranslation

def printUniChrs(s):
    tmpst = ''
    for x in s:
        if ord(x) < 128:
            if not ord(x) in [38]:  # 38 = ampersand
#                print x,
                tmpst += x
#            else:
#                print ' ',
        else:
#            print ' ',
            tmpst += '\\x%x' % ord(x)
#	print ord(x), "%x" % ord(x)
    print
    print tmpst


languageOptions = ['English', 'Unicode', 'Chinese', 'Arabic']

clNum = 3

currentLanguage = languageOptions[clNum]


tmpApp = wx.PySimpleApp()

print
print "Current Language = " + languageOptions[clNum]
print

if currentLanguage == 'English':
    presLan_en.install()  # Adding "unicode=1" here would eliminiate the need to declare translations strings as unicode() objects!!
    lang = wx.LANGUAGE_ENGLISH

elif currentLanguage == 'Unicode':
    # French
    dir = os.path.join('C:\\Users\\DavidWoods\\Documents\\Python\\Transana\\transana\\transana\\src', 'locale', 'fr', 'LC_MESSAGES', 'Transana.mo')
    if os.path.exists(dir):
        presLan_fr = gettext.translation('Transana', 'locale', languages=['fr']) # French
    lang = wx.LANGUAGE_FRENCH
    presLan_fr.install()  # Adding "unicode=1" here would eliminiate the need to declare translations strings as unicode() objects!!
    words = ['&French']

elif currentLanguage == 'Chinese':
    # Chinese
    dir = os.path.join('C:\\Users\\DavidWoods\\Documents\\Python\\Transana\\transana\\transana\\src', 'locale', 'zh', 'LC_MESSAGES', 'Transana.mo')
    if os.path.exists(dir):
        presLan_zh = gettext.translation('Transana', 'locale', languages=['zh']) # Chinese
    lang = wx.LANGUAGE_CHINESE
    presLan_zh.install()  # Adding "unicode=1" here would eliminiate the need to declare translations strings as unicode() objects!!
    words = ['&Chinese - Simplified']

elif currentLanguage == 'Arabic':
    # Arabic
    dir = os.path.join('C:\\Users\\DavidWoods\\Documents\\Python\\Transana\\transana\\transana\\src', 'locale', 'ar', 'LC_MESSAGES', 'Transana.mo')
    if os.path.exists(dir):
        presLan_ar = gettext.translation('Transana', 'locale', languages=['ar']) # Arabic
    lang = wx.LANGUAGE_ARABIC
    presLan_ar.install()  # Adding "unicode=1" here would eliminiate the need to declare translations strings as unicode() objects!!
    words = ['&Arabic']

# This provides localization for wxPython
locale = wx.Locale(lang) # , wx.LOCALE_LOAD_DEFAULT | wx.LOCALE_CONV_ENCODING)

# Add the Transana catalog to the Locale
locale.AddCatalog("Transana")

words += [
#          'Keyword Group', "Keyword", 'Definition',
#          'Series', 'Owner', 'Comment',
#          'Episode', 'Comment',
#          'Transcript', 'Transcriber', 'Comment',
#          'Collection', 'Owner', 'Comment',
#          'Clip', 'Comment',
#          'Snapshot', 'Comment',
#           'Note', 'Comment', 'Note Taker', 'Note Text:',
#           'Series', 'Episode', 'Transcript', 'Collection', 'Clip', 'Snapshot'
           'Title', 'Creator', 'Subject', 'Description', 'Publisher', 'Contributor', 'Media Type', 'Format', 'Source', 'Language',
           'Relation', 'Coverage', 'Rights'
         ]

for word in words:
    print word

    tmp = _(word)


    tmp = tmp.decode('utf8')

#    print printUniChrs(tmp)

#    print

    tmp = tmp.encode('utf8')

    print printUniChrs(tmp)
    print
    print ' - - - - - - - - - - - - - - - - - - - - - - - - - - - -'
    print

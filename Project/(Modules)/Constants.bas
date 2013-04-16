Attribute VB_Name = "modConstants"
Option Explicit


'--- Charset Constants
Public Const VOWELS As String = "eaoiu"   'In order of frequency
Public Const CONSONANTS As String = "tnshrdlcmfgypwbvkxjqz"   'In order of frequency
Public Const SYMBOLS    As String = "! @ # % $ ^ & * #lpa# #rpa# #lbr# #rbr# : ' / ` ~ * - < > #pls# = _ #pip# #sla# #sla# . . , , ; ; ? ? #lba# #rba#"
Public Const KEYBOARDSYMBOLS As String = "`~!@#$%^&*()-_=+]}[{\|'"";:/?.>,< "
Public Const NONKEYBOARDSYMBOLS As String = " ¡¢£¤¥¦§¨©ª«¬­®¯°±²³´µ¶·¸¹º»¼½¾¿ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖ×ØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõö÷øùúûüýþÿ"
Public Const NONPRINTABLE = "" & " " & vbCrLf
Public Const SENTENCEPUNCTUATION As String = "!;:?.,"
Public Const ENDPUNCTUATION As String = "! ! ! ! . . . . . . . . . . . . . . . ... ... ? ? ? ? ? ? ?"
Public Const LETTERS As String = "etaoinshrdlucmfgypwbvkxjqz"   'In order of frequency
Public Const UCASELETTERS As String = "ABCDEFGHIJKLMNOPQRSTUVWYZ"
Public Const LCASELETTERS As String = "abcdefghijklmnopqrstuvwxyz"
Public Const NUMBERS    As String = "0123456789"
Public Const SMILEYS    As String = ":) :( :-) :-( :D :0 ;-) ;) :/ 8-) 8-( :-D :-0 :-p :^)"
Public Const VOWELS2    As String = "a a a a a a a a a e e e e e e e e e e e i i i u u o o ay ea ee ia io oa oi oo er on re he ha in es io ou "
Public Const CONSONANTS2 As String = "b b c d d d f g j k m m m n n p p qu r r r s s s s t t t t v w x z z th st sh ph ch th sh for has tis men"   'can appear anywhere in a word
Public Const CONSONANTS3 As String = "nd rt dd zz rg ng tt ss mm nn pp nt nc nl ft"   'A word cannot start with any of these
Public Const CONFUSING    As String = "1 l 0 O o __ ___ 5 z 2 Z i I"
Public Const KEYBOARD    As String = "1234567890`~!@#$%^&*()-_=+]}[{\|'"";:/?.>,<abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

Public Const NUMROW As String = "1234567890"   'Left to right
Public Const NUMROWFULL As String = "1234567890`~!@#$%^&*()_-+="
Public Const ROW1 As String = "QWERTYUIOP"    'Left to right
Public Const ROW1FULL As String = "QWERTYUIOP{[}]|\"
Public Const ROW2 As String = "ASDFGHJKL"    'Left to right
Public Const ROW2FULL As String = "ASDFGHJKL;:'"""
Public Const ROW3 As String = "ZXCVBNM"    'Left to right
Public Const ROW3FULL As String = "ZXCVBNM,<.>/?"
Public Const LEFTHAND As String = "qwertasdfgzxcvb"
Public Const RIGHTHAND As String = "yuiophjknm"
Public Const DIGRAPHS As String = "th er on an re he in ed nd ha at en es of or nt ea ti to it st io le is ou ar as de rt ve" 'In order of frequency
Public Const TRIGRAPHS As String = "the and tha ent ion tio for nde has nce edt tis oft sth men" 'In order of frequency
Public Const INITIALLETTERS As String = "T0AWBCDSFMRHIYEGLNPUJK " 'In order of frequency
Public Const TWOLETTERWORDS As String = "of to in it is be as at so we he by or on do if me my up an go no us am" 'In order of frequency
Public Const THREELETTERWORDS As String = "the and for are but not you all any can had her was one our out day get has him his how man new now old see two way who boy did its let put say she too use" 'In order of frequency
Public Const LONGMONTHS As String = "January February March April May June July August September October November December"
Public Const SHORTMONTHS As String = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec"
Public Const LONGDAYS As String = "Monday Tuesday Wednesday Thursday Friday Saturday Sunday"
Public Const SHORTDAYS As String = "Mon Tue Wed Thu Fri Sat Sun"

Public Const SP0 = "[--If you are reading this, good for you.--]"

Private Const SP1 = "3C2C2E2EFFF05B6F0492F63B6751200894F554C4A8DB629335EFB1D269D4CCC42FAEFF2A0C03E5A90E854521F930E61A937B81EAFB6C7CFA4DC3D2C526726801184EA8019529E67403B39D3E5023C9874C514FD268992B3A0B815EB290DA593C8D4936385967"
Private Const SP2 = "3C2C2E2EFFF00BB86E835A91BBFCBF91FF7EE128CFE8A73A61BAAA1F9030D3FD978AE20051273EE750F404E259B92C47234682BFC55F82E36BA2ABF28AFF3FB7DB608E647AA34F52718001A84DC74083B86DE9843BA93D704FE168C1127F925155C1E8C92878"
Private Const SP3 = "3C2C2E2EFFF0935DA570B63AF5C5B5170724DE3581B75FCC48B16D1F0A789F7200C110A0C34E8BB76659533A7680E1272122DE500C65F41A3934308A9EBEB95EF4FCE7F4428AC7357C27C0CF4825A9210925899751B0901345BB9464D3CAAF88779229A36D10"
Private Const SP4 = "3C2C2E2EFFF05645DDA5269E5FD38A6A46558AB785057F5DB09F6AACCBF9A056991C804BEBA156E54AC43FAA17DF6CF8A5824BB99E01E6E21F13754EEE2CC18AB5150B9B63FB2B29498B3D71DB9B1FEA0F80291464A63E00FE3E3CDAEE9C663C32A55065DF4E"


'------------ Password Warnings & Errors
Public Const PASSWORD_WARNING = ERRBASE + 4096
Public Const BEGINS_OR_ENDS_WITH_SPACE = PASSWORD_WARNING + 1      '  0
Public Const PASSWORD_TOO_SHORT = PASSWORD_WARNING + 2
'Public Const SHORT_AND_ENDS_WITH_0 = PASSWORD_WARNING + 4          ' -1
Public Const ALL_NUMERIC_PASSWORD = PASSWORD_WARNING + 8         ' -1
Public Const TRIVIAL_WORD = PASSWORD_WARNING + 16                  ' -1
Public Const TRIVIAL_SEQUENCE = PASSWORD_WARNING + 32              ' -1
Public Const TRIVIAL_REPETITION = PASSWORD_WARNING + 64            ' -2
Public Const TOO_MANY_SINGLE_CHARSET = PASSWORD_WARNING + 128      ' -1
Public Const WORDLIST_MATCH = PASSWORD_WARNING + 256               ' -4
Public Const PARTIAL_WORDLIST_MATCH = PASSWORD_WARNING + 512       ' -3
Public Const REVERSED_WORDLIST_MATCH = PASSWORD_WARNING + 1024   ' -2
Public Const NOT_ENOUGH_CHARSETS = PASSWORD_WARNING + 2048         ' -1
Public Const TOO_MANY_SINGLE_CHARACTER = PASSWORD_WARNING + 4096   ' -1
Public Const SIMILAR_WORDLIST_MATCH = PASSWORD_WARNING + 8192      ' -2
Public Const MISSING_REQUIRED_CHAR = PASSWORD_WARNING + 16384
Public Const CONTAINS_RESTRICTED_CHAR = PASSWORD_WARNING + 32768
Public Const CONTAINS_USER_DATA = PASSWORD_WARNING + 65536         ' -3
Public Const COMMON_PASSWORD_PATTERN = PASSWORD_WARNING + 131072   ' -2
Public Const BEGINS_WITH_RESTRICTED_CHAR = PASSWORD_WARNING + 262144
Public Const ENDS_WITH_RESTRICTED_CHAR = PASSWORD_WARNING + 524288


'------------ Password Warning Descriptions (EN-US)
Public Const BEGINS_OR_ENDS_WITH_SPACE_DESC = "Password ends or starts with a space, which might cause problems on some systems"
Public Const PASSWORD_TOO_SHORT_DESC = "Password is too short"
'Public Const SHORT_AND_ENDS_WITH_0_DESC = "Password is too short and ends with a 0"
Public Const ALL_NUMERIC_PASSWORD_DESC = "Password is all numberic"
Public Const TRIVIAL_WORD_DESC = "Password contains common word, phrase, or pattern"
Public Const TRIVIAL_SEQUENCE_DESC = "Password contains common character sequence"
Public Const TRIVIAL_REPETITION_DESC = "Password contains trivial repetition"
Public Const TOO_MANY_SINGLE_CHARSET_DESC = "Password contains too many similar characters"
Public Const WORDLIST_MATCH_DESC = "Password matches word in wordlist"
Public Const PARTIAL_WORDLIST_MATCH_DESC = "Password partially matches word in wordlist"
Public Const REVERSED_WORDLIST_MATCH_DESC = "Password reversed matches word in wordlist"
Public Const NOT_ENOUGH_CHARSETS_DESC = "Password does not contain enough different character sets"
Public Const TOO_MANY_SINGLE_CHARACTER_DESC = "Password contains too many of the same character"
Public Const SIMILAR_WORDLIST_MATCH_DESC = "Password is too similar to a word in the wordlist"
Public Const MISSING_REQUIRED_CHAR_DESC = "Password does not contain a required character"
Public Const CONTAINS_RESTRICTED_CHAR_DESC = "Password contains a restricted character"
Public Const CONTAINS_USER_DATA_DESC = "Password contains user information"
Public Const COMMON_PASSWORD_PATTERN_DESC = "Password follows too common a pattern"


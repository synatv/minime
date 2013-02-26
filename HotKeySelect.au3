;===============================================================================
; Name:			HotKeySelect
; Author:		Prog@ndy (Modified by Tijs Hendriks - http://www.saphua.com/)
; Description:	Creates an input box that allows selection of hotkeys
;===============================================================================

#include <WinAPI.au3>

global const $HKM_GETHOTKEY = $WM_USER + 2
global const $HOTKEYF_ALT = 0x4
global const $HOTKEYF_CONTROL = 0x2
global const $HOTKEYF_SHIFT = 0x1
global const $HOTKEYF_EXT = 0x80

global $VK_HotKey_Mapping[65][2]
_CreateVK()

func _GetHotKey($hotkeyControl)
	if not IsHWnd($hotkeyControl) then return -1
	local $Chr = ""
	local $hk = _SendMessage($hotkeyControl, $HKM_GETHOTKEY, 0, 0)
	local $loByte = BitAND(_WinAPI_LoWord($hk), 0xFF)
	local $HiByte = BitShift(_WinAPI_LoWord($hk), 8)
	$hk = $loByte
	for $i = 0 To 64
		if $VK_HotKey_Mapping[$i][0] = $hk then
			$Chr = $VK_HotKey_Mapping[$i][1]
			exitloop
		endif
	next
	if $Chr <> "" then return $Chr
	$Chr = DllCall("user32.dll", "int", "MapVirtualKey", "int", $hk, "int", 2)
	$Chr = StringLower(Chr(BitAND($Chr[0], 0xFFFF)))
	if $Chr = "!" or $Chr = "^" or $Chr = "+" or $Chr = "{" or $Chr = "}" or $Chr = "#" then
		$Chr = "{" & $Chr & "}"
	endif
	if BitAND($HiByte, $HOTKEYF_SHIFT) = $HOTKEYF_SHIFT then $Chr = "+" & $Chr
	if BitAND($HiByte, $HOTKEYF_ALT) = $HOTKEYF_ALT then $Chr = "!" & $Chr
	if BitAND($HiByte, $HOTKEYF_CONTROL) = $HOTKEYF_CONTROL then $Chr = "^" & $Chr
	if BitAND($HiByte, $HOTKEYF_EXT) = $HOTKEYF_EXT then $Chr = "#" & $Chr
	return $Chr
EndFunc ;==>_GetHotKey

Func _CreateVK()
	;~ Global Const $VK_CTRL_BREAK = '03'
	$VK_HotKey_Mapping[0][0] = 0x03
	$VK_HotKey_Mapping[0][1] = "{BREAK}"
	;~ Global Const $VK_BACK = '08'
	$VK_HotKey_Mapping[1][0] = 0x09
	$VK_HotKey_Mapping[1][1] = "{TAB}"
	;~ Global Const $VK_TAB = '09'
	;~ Global Const $VK_CLEAR = '0C'
	$VK_HotKey_Mapping[2][0] = 0x0C
	$VK_HotKey_Mapping[2][1] = "{CLEAR}"
	;~ Global Const $VK_ENTER = '0D'
	$VK_HotKey_Mapping[3][0] = 0x0D
	$VK_HotKey_Mapping[3][1] = "{ENTER}"
	;~ Global Const $VK_SHIFT = 10
	$VK_HotKey_Mapping[4][0] = 0x10
	$VK_HotKey_Mapping[4][1] = "+"
	;~ Global Const $VK_CTRL = 11
	$VK_HotKey_Mapping[5][0] = 0x11
	$VK_HotKey_Mapping[5][1] = "^"
	;~ Global Const $VK_ALT = 12
	$VK_HotKey_Mapping[6][0] = 0x12
	$VK_HotKey_Mapping[6][1] = "!"
	;~ Global Const $VK_PAUSE = 13
	$VK_HotKey_Mapping[7][0] = 0x13
	$VK_HotKey_Mapping[7][1] = "{PAUSE}"
	;~ Global Const $VK_CAPS = 14
	$VK_HotKey_Mapping[8][0] = 0x14
	$VK_HotKey_Mapping[8][1] = "{CAPSLOCK}"
	;~ Global Const $VK_ESC = '1B'
	$VK_HotKey_Mapping[9][0] = 0x1B
	$VK_HotKey_Mapping[9][1] = "{ESC}"
	;~ Global Const $VK_SPACE = 20
	$VK_HotKey_Mapping[10][0] = 0x20
	$VK_HotKey_Mapping[10][1] = "{SPACE}"
	;~ Global Const $VK_PAGE_UP = 21
	$VK_HotKey_Mapping[11][0] = 0x21
	$VK_HotKey_Mapping[11][1] = "{PGUP}"
	;~ Global Const $VK_PADE_DOWN = 22
	$VK_HotKey_Mapping[12][0] = 0x22
	$VK_HotKey_Mapping[12][1] = "{PGDN}"
	;~ Global Const $VK_END = 23
	$VK_HotKey_Mapping[13][0] = 0x23
	$VK_HotKey_Mapping[13][1] = "{END}"
	;~ Global Const $VK_HOME = 24
	$VK_HotKey_Mapping[14][0] = 0x24
	$VK_HotKey_Mapping[14][1] = "{HOME}"
	;~ Global Const $VK_LEFT = 25
	$VK_HotKey_Mapping[15][0] = 0x25
	$VK_HotKey_Mapping[15][1] = "{LEFT}"
	;~ Global Const $VK_UP = 26
	$VK_HotKey_Mapping[16][0] = 0x26
	$VK_HotKey_Mapping[16][1] = "{UP}"
	;~ Global Const $VK_RIGHT = 27
	$VK_HotKey_Mapping[17][0] = 0x27
	$VK_HotKey_Mapping[17][1] = "{RIGHT}"
	;~ Global Const $VK_DOWN = 28
	$VK_HotKey_Mapping[18][0] = 0x28
	$VK_HotKey_Mapping[18][1] = "{DOWN}"
	;~ Global Const $VK_SELECT = 29
	$VK_HotKey_Mapping[19][0] = 0x29
	$VK_HotKey_Mapping[19][1] = "{SELECT}"
	;~ Global Const $VK_PRINT = '2A'
	$VK_HotKey_Mapping[20][0] = 0x2A
	$VK_HotKey_Mapping[20][1] = "{PRINTSCREEN}"
	;~ Global Const $VK_EXECUTE = '2B'
	;~ $VK_HotKey_Mapping[1][0]=0x20
	;~ $VK_HotKey_Mapping[1][1]="{SPACE}"
	;~ Global Const $VK_PRINT_SCR = '2C'
	$VK_HotKey_Mapping[21][0] = 0x2C
	$VK_HotKey_Mapping[21][1] = "{PRINTSCREEN}"
	;~ Global Const $VK_INS = '2D'
	$VK_HotKey_Mapping[22][0] = 0x2D
	$VK_HotKey_Mapping[22][1] = "{INS}"
	;~ Global Const $VK_DEL = '2E'
	$VK_HotKey_Mapping[23][0] = 0x2E
	$VK_HotKey_Mapping[23][1] = "{DEL}"
	;~ Global Const $VK_HELP = '2F'
	$VK_HotKey_Mapping[24][0] = 0x2F
	$VK_HotKey_Mapping[24][1] = "{F1}"
	;~ Global Const $VK_L_WIN = '5B'
	$VK_HotKey_Mapping[25][0] = 0x5B
	$VK_HotKey_Mapping[25][1] = "#"
	;~ Global Const $VK_R_WIN = '5C'
	$VK_HotKey_Mapping[26][0] = 0x5C
	$VK_HotKey_Mapping[26][1] = "#"
	;~ Global Const $VK_APP = '5D'
	$VK_HotKey_Mapping[27][0] = 0x5D
	$VK_HotKey_Mapping[27][1] = "{APPSKEY}"
	;~ Global Const $VK_NUMPAD0 = 60
	$VK_HotKey_Mapping[28][0] = 0x60
	$VK_HotKey_Mapping[28][1] = "{NUMPAD0}"
	;~ Global Const $VK_NUMPAD1 = 61
	$VK_HotKey_Mapping[29][0] = 0x61
	$VK_HotKey_Mapping[29][1] = "{NUMPAD1}"
	;~ Global Const $VK_NUMPAD2 = 62
	$VK_HotKey_Mapping[30][0] = 0x62
	$VK_HotKey_Mapping[30][1] = "{NUMPAD2}"
	;~ Global Const $VK_NUMPAD3 = 63
	$VK_HotKey_Mapping[31][0] = 0x63
	$VK_HotKey_Mapping[31][1] = "{NUMPAD3}"
	;~ Global Const $VK_NUMPAD4 = 64
	$VK_HotKey_Mapping[32][0] = 0x64
	$VK_HotKey_Mapping[32][1] = "{NUMPAD4}"
	;~ Global Const $VK_NUMPAD5 = 65
	$VK_HotKey_Mapping[33][0] = 0x65
	$VK_HotKey_Mapping[33][1] = "{NUMPAD5}"
	;~ Global Const $VK_NUMPAD6 = 66
	$VK_HotKey_Mapping[34][0] = 0x66
	$VK_HotKey_Mapping[34][1] = "{NUMPAD6}"
	;~ Global Const $VK_NUMPAD7 = 67
	$VK_HotKey_Mapping[35][0] = 0x67
	$VK_HotKey_Mapping[35][1] = "{NUMPAD7}"
	;~ Global Const $VK_NUMPAD8 = 68
	$VK_HotKey_Mapping[36][0] = 0x68
	$VK_HotKey_Mapping[36][1] = "{NUMPAD8}"
	;~ Global Const $VK_NUMPAD9 = 69
	$VK_HotKey_Mapping[37][0] = 0x69
	$VK_HotKey_Mapping[37][1] = "{NUMPAD9}"
	;~ Global Const $VK_MULTIPLY = '6A'
	$VK_HotKey_Mapping[38][0] = 0x6A
	$VK_HotKey_Mapping[38][1] = "{NUMPADMULT}"
	;~ Global Const $VK_ADD = '6B'
	$VK_HotKey_Mapping[39][0] = 0x6B
	$VK_HotKey_Mapping[39][1] = "{NUMPADADD}"
	;~ Global Const $VK_SEPERATOR = '6C'
	$VK_HotKey_Mapping[40][0] = 0x6C
	$VK_HotKey_Mapping[40][1] = "{NUMPADENTER}"
	;~ Global Const $VK_SUBSTRACT = '6D'
	$VK_HotKey_Mapping[41][0] = 0x6D
	$VK_HotKey_Mapping[41][1] = "{NUMPADSUB}"
	;~ Global Const $VK_DECIMAL = '6E'
	$VK_HotKey_Mapping[42][0] = 0x6E
	$VK_HotKey_Mapping[42][1] = "{NUMPADDOT}"
	;~ Global Const $VK_DIVIDE = '6F'
	$VK_HotKey_Mapping[43][0] = 0x6F
	$VK_HotKey_Mapping[43][1] = "{NUMPADDIV}"
	;~ Global Const $VK_F1 = 70
	$VK_HotKey_Mapping[44][0] = 0x70
	$VK_HotKey_Mapping[44][1] = "{F1}"
	;~ Global Const $VK_F2 = 71
	$VK_HotKey_Mapping[45][0] = 0x71
	$VK_HotKey_Mapping[45][1] = "{F2}"
	;~ Global Const $VK_F3 = 72
	$VK_HotKey_Mapping[46][0] = 0x72
	$VK_HotKey_Mapping[46][1] = "{F3}"
	;~ Global Const $VK_F4 = 73
	$VK_HotKey_Mapping[47][0] = 0x73
	$VK_HotKey_Mapping[47][1] = "{F4}"
	;~ Global Const $VK_F5 = 74
	$VK_HotKey_Mapping[48][0] = 0x74
	$VK_HotKey_Mapping[48][1] = "{F5}"
	;~ Global Const $VK_F6 = 75
	$VK_HotKey_Mapping[49][0] = 0x75
	$VK_HotKey_Mapping[49][1] = "{F6}"
	;~ Global Const $VK_F7 = 76
	$VK_HotKey_Mapping[50][0] = 0x76
	$VK_HotKey_Mapping[50][1] = "{F7}"
	;~ Global Const $VK_F8 = 77
	$VK_HotKey_Mapping[51][0] = 0x77
	$VK_HotKey_Mapping[51][1] = "{F8}"
	;~ Global Const $VK_F9 = 78
	$VK_HotKey_Mapping[52][0] = 0x78
	$VK_HotKey_Mapping[52][1] = "{F9}"
	;~ Global Const $VK_F10 = 79
	$VK_HotKey_Mapping[53][0] = 0x79
	$VK_HotKey_Mapping[53][1] = "{F10}"
	;~ Global Const $VK_F11 = '7A'
	$VK_HotKey_Mapping[54][0] = 0x7A
	$VK_HotKey_Mapping[54][1] = "{F11}"
	;~ Global Const $VK_F12 = '7B'
	$VK_HotKey_Mapping[55][0] = 0x7B
	$VK_HotKey_Mapping[55][1] = "{F12}"
	;~ Global Const $VK_F13 = '7C'
	;~ Global Const $VK_F14 = '7D'
	;~ Global Const $VK_F15 = '7E'
	;~ Global Const $VK_F16 = '7F'
	;~ Global Const $VK_F17 = '80H'
	;~ Global Const $VK_F18 = '81H'
	;~ Global Const $VK_F19 = '82H'
	;~ Global Const $VK_F20 = '83H'
	;~ Global Const $VK_F21 = '84H'
	;~ Global Const $VK_F22 = '85H'
	;~ Global Const $VK_F23 = '86H'
	;~ Global Const $VK_F24 = '87H'
	;~ Global Const $VK_NUMLOCK = 90
	$VK_HotKey_Mapping[56][0] = 0x90
	$VK_HotKey_Mapping[56][1] = "{NUMLOCK}"
	;~ Global Const $VK_SCROLL_LOCK = 91
	$VK_HotKey_Mapping[57][0] = 0x91
	$VK_HotKey_Mapping[57][1] = "{SCROLLLOCK}"
	;~ Global Const $VK_L_SHIFT = 'A0'
	$VK_HotKey_Mapping[58][0] = 0xA0
	$VK_HotKey_Mapping[58][1] = "+"
	;~ Global Const $VK_R_SHIFT = 'A1'
	$VK_HotKey_Mapping[59][0] = 0xA1
	$VK_HotKey_Mapping[59][1] = "+"
	;~ Global Const $VK_L_CTRL = 'A2'
	$VK_HotKey_Mapping[60][0] = 0xA2
	$VK_HotKey_Mapping[60][1] = "^"
	;~ Global Const $VK_R_CTRL = 'A3'
	$VK_HotKey_Mapping[61][0] = 0xA3
	$VK_HotKey_Mapping[61][1] = "^"
	;~ Global Const $VK_L_MENU = 'A4'
	$VK_HotKey_Mapping[62][0] = 0xA4
	$VK_HotKey_Mapping[62][1] = "{APPSKEY}"
	;~ Global Const $VK_R_MENU = 'A5'
	$VK_HotKey_Mapping[63][0] = 0xA5
	$VK_HotKey_Mapping[63][1] = "{APPSKEY}"
	;~ Global Const $VK_PLAY = 'FA'
	$VK_HotKey_Mapping[64][0] = 0xFA
	$VK_HotKey_Mapping[64][1] = "{MEDIA_PLAY_PAUSE}"
	;~ Global Const $VK_ZOOM = 'FB'
	;~ Global Const $VK_OFF = 'DF'
	;~ Global Const $VK_COMMA = 'BC'
	;~ Global Const $VK_POINT = 'BE'
	;~ Global Const $VK_PERIOD = 'BE'
	;~ Global Const $VK_PLUS = 'BB'
	;~ Global Const $VK_MINUS = 'BD'
	;other:
	;~ Global Const $VK_COLON = 'BA' ;==> :;
	;~ Global Const $VK_SLASH = 'BF' ;==> /?
	;~ Global Const $VK_TILDE = 'C0' ;==> `~
	;~ Global Const $VK_OPEN_BRACKET = 'DB' ;==> [{
	;~ Global Const $VK_CLOSE_BRACKET = 'DD' ;==> ]}
	;~ Global Const $VK_BACK_SLASH = 'DC' ;==> \|
	;~ Global Const $VK_QUOTATION = 'DE' ;==> '"
EndFunc ;==>_CreateVK
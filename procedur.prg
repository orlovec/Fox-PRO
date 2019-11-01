* PROCEDUR.PRG
* ����������� ����

* �������� ����� � ���������� ������
FUNCTION MD_EX
LPARAMETERS _Path as Character
LOCAL loErr as Exception, lcMsg as Character
lcMsg = []
TRY 
	MD (_Path)
CATCH TO loErr
	lcMsg = loErr.Message
ENDTRY

RETURN (lcMsg)
ENDFUNC



* ��������� ������� ��� �����
*
PROCEDURE new_filetime(lFileName, lFileTime)  
 * lFileName ���� � �����, ������ chr  
 * lFileTime ����� �����, ������ datetime   
    
  LOCAL LFT, UTC, buf, hnd, f_date, f_time  
 *~*	DECLARE INTEGER GetTimeZoneInformation IN WIN32API string @struct1  
 *~*	struct1=SPACE(172)  
 *~*	is_daylight = GetTimeZoneInformation(@struct1)  
 *~*	?(CTOBIN(CHR(0x7F)+SUBSTR(struct1,1,1))+IIF(is_daylight=2,CTOBIN(CHR(0x7F)+SUBSTR(struct1,169,1)),0))/60  
    
  DECLARE INTEGER DosDateTimeToFileTime IN WIN32API integer dos_date, integer dos_time, string @filetime  
  DECLARE INTEGER SetFileTime IN WIN32API integer handle, string @creation, string @access, string @wright  
  DECLARE INTEGER LocalFileTimeToFileTime IN WIN32API string @filetime, string @UTC  
  DECLARE INTEGER OpenFile IN Win32API string FileName, string @Buffer, integer AccesType    
  DECLARE CloseHandle in Win32API INTEGER Handle    
    
  LFT=SPACE(8)  
  UTC=SPACE(8)  
  buf=space(150)  
    
  f_date=BITOR(DAY(lFileTime), BITLSHIFT(MONTH(lFileTime),5), BITLSHIFT(YEAR(lFileTime)-1980,9))  
  f_time=BITOR(SEC(lFileTime)/2, BITLSHIFT(MINUTE(lFileTime),5),  BITLSHIFT(HOUR(lFileTime),11))  
  DosDateTimeToFileTime(f_date, f_time, @LFT)  
  LocalFileTimeToFileTime(@LFT, @UTC)  
  hnd=OpenFile(lFileName, @buf, 1)    
  SetFileTime(hnd, null, null, UTC)    
  CloseHandle(hnd)    
    
  CLEAR DLLS DosDateTimeToFileTime, SetFileTime, LocalFileTimeToFileTime, OpenFile, CloseHandle   
  ENDPROC


FUNCTION HTMLColor  
  Lparameters m.nColor as Integer
  Local s  
  s=Transform(m.nColor,'@0')  
  Return '#'+Substr(s,9,2)+Substr(s,7,2)+Substr(s,5,2)
  
*!*	* ���������� ������������
FUNCTION StopPlayFile(cFile)
IF !FILE(cFile)
	RETURN(.F.)
ENDIF
LOCAL cDll, cStr
DO CASE
CASE OS() = "Windows 5.00"
	cDll = GETENV("windir")+"\SYSTEM32\winmm.dll"
CASE OS() = "Windows 5.01"
	cDll = GETENV("windir")+"\SYSTEM32\winmm.dll"
OTHERWISE
	cDll = GETENV("windir")+"\SYSTEM\winmm.dll"
ENDCASE
IF !FILE(cDll)
	RETURN (.F.)
ENDIF

cStr = [close ]+cFile
Declare Long mciExecute IN (cDll) String cFile

mciExecute(cStr)

CLEAR DLLS "mciExecute"
RETURN (.T.)
ENDFUNC

*!*	* ������ ������������
FUNCTION PlayAnyFile(cFile)
IF !FILE(cFile)
	RETURN(.F.)
ENDIF
LOCAL cDll, cStr
DO CASE
CASE OS() = "Windows 5.00"
	cDll = GETENV("windir")+"\SYSTEM32\winmm.dll"
CASE OS() = "Windows 5.01"
	cDll = GETENV("windir")+"\SYSTEM32\winmm.dll"
OTHERWISE
	cDll = GETENV("windir")+"\SYSTEM\winmm.dll"
ENDCASE
IF !FILE(cDll)
	RETURN (.F.)
ENDIF

cStr = [play ]+cFile
Declare Long mciExecute IN (cDll) String cFile

mciExecute(cStr)

*!*	* ���������� ������������
*!*	mciExecute("close pinmidi.wav")
CLEAR DLLS "mciExecute"
RETURN (.T.)
ENDFUNC

*
* ������ ������ �� ��������� �����: 2013,2014,2015
* ���������: liNum_Start-��������� �����, liNum_End-�������� �����, lbOrder-�������: .T.-� ������������ (2013,2014,2015), .F.-� ��������� (2015,2014,2013)
* ����: 08-08-2012
*
FUNCTION Number2String
LPARAMETERS liNum_Start as Number, liNum_End as Number, lbOrder as Boolean
LOCAL liI as Number
IF lbOrder
	FOR liI=liNum_Start TO liNum_End
		
	ENDFOR
ELSE
ENDIF
ENDFUNC
*

FUNCTION DayOfWeek(dT)
_n = DOW(dT,2)
DO CASE
CASE _n = 1
	_s = "�����������"
CASE _n = 2
	_s = "�������"
CASE _n = 3
	_s = "�����"
CASE _n = 4
	_s = "�������"
CASE _n = 5
	_s = "�������"
CASE _n = 6
	_s = "�������"
CASE _n = 7
	_s = "�����������"
ENDCASE
RETURN (_s)
ENDFUNC

*
FUNCTION DataRus(dT)
IF PCOUNT()=0
	dT = DATE()
ENDIF
LOCAL cStr as String
LOCAL ARRAY cM[12]
DIMENSION cM[12]
cM[1] = "������"
cM[2] = "�������"
cM[3] = "�����"
cM[4] = "������"
cM[5] = "���"
cM[6] = "����"
cM[7] = "����"
cM[8] = "�������"
cM[9] = "��������"
cM[10] = "�������"
cM[11] = "������"
cM[12] = "�������"
cStr = lzerro(DAY(dT),2)+' '+cM[MONTH(dT)]+' '+STR(YEAR(dT),4)+' �.'
RETURN (cStr)
ENDFUNC

* �������������� �������� ������ � ����� ������
FUNCTION MonthFromStr(lcS)
LOCAL _s
_s = 0
DO CASE
CASE ATC("�����",lcS)>0
	_s = 1
CASE ATC("������",lcS)>0
	_s = 2
CASE ATC("����",lcS)>0
	_s = 3
CASE ATC("������",lcS)>0
	_s = 4
CASE (ATC("���",lcS)>0 OR ATC("���",lcS)>0)
	_s = 5
CASE ATC("����",lcS)>0
	_s = 6
CASE ATC("����",lcS)>0
	_s = 7
CASE ATC("������",lcS)>0
	_s = 8
CASE ATC("��������",lcS)>0
	_s = 9
CASE ATC("������",lcS)>0
	_s = 10
CASE ATC("�����",lcS)>0
	_s = 11
CASE ATC("������",lcS)>0
	_s = 12
ENDCASE
RETURN (_s)
ENDFUNC

* �������������� ������ ������ � �������� � �������������� ���������
FUNCTION RusMonth_EX(nM, nPos)
IF PCOUNT() = 0
	nM = MONTH(DATE())
	nPos = 1
ENDIF
IF PCOUNT()=1
	nPos = 1
ENDIF
IF nPos<=0 OR nPos>2
	nPos = 1
ENDIF
LOCAL cM
DIMENSION cM[12,2]
cM[1,1] = "������"
cM[1,2] = "������"

cM[2,1] = "�������"
cM[2,2] = "�������"

cM[3,1] = "����"
cM[3,2] = "�����"

cM[4,1] = "������"
cM[4,2] = "������"

cM[5,1] = "���"
cM[5,2] = "���"

cM[6,1] = "����"
cM[6,2] = "����"

cM[7,1] = "����"
cM[7,2] = "����"

cM[8,1] = "������"
cM[8,2] = "�������"

cM[9,1] = "��������"
cM[9,2] = "��������"

cM[10,1] = "�������"
cM[10,2] = "�������"

cM[11,1] = "������"
cM[11,2] = "������"

cM[12,1] = "�������"
cM[12,2] = "�������"
RETURN (cM[nM,nPos])
ENDFUNC


FUNCTION RusMonth(nM)
IF PCOUNT() = 0
	nM = MONTH(DATE())
ENDIF
LOCAL cM
DIMENSION cM[12]
cM[1] = "������"
cM[2] = "�������"
cM[3] = "����"
cM[4] = "������"
cM[5] = "���"
cM[6] = "����"
cM[7] = "����"
cM[8] = "������"
cM[9] = "��������"
cM[10] = "�������"
cM[11] = "������"
cM[12] = "�������"
RETURN (cM[nM])
ENDFUNC

Function MYToN
parameters MM_, YY_
if parameters() < 2
	MM_ = month(date())
	YY_ = year(date())
endif
if MM_ < 1
	return YY_
endif
if YY_ < 1
	return MM_
endif
return (YY_-1)*12+MM_

Function NToMY
parameters NN_, KOD_
if parameters() < 2
	return 0
endif
if KOD_ = 0
	return iif(mod(NN_, 12) = 0, 12, mod(NN_,12))
else
	return int((NN_-1)/12)+1
endif
return 0

* FUNNCTION CPath()
* ���������: ����
* ���� ��� ��� �� ������������� ��� �������
*
Function CPath
parameters _Path
if parameters()=0
	return .f.
endif
private _home, _errProc
_errProc = on("ERROR")
on error do CPathError with error()
_home=sys(5)+sys(2003)
chdir &_Path

if !_Path $ sys(2003)
	mkdir &_Path
endif
set default to (_home)
on error &_errProc
return .t.

Function CPathError
parameters _r
return


* ������� FILEDBF
* ���������: ���, �����, ��������
* ������������ �������� ���� 9704DBF
*
FUNCTION FileDbf
PARAMETERS _y,_m,_pr
do case
case parameters()=0
	_y=year(date())
	_m=month(date())
    _pr=''
case parameters()=1
    return right(lzerro(str(_y,4)),4)
case parameters()=2
	if _m=0
		return right(lzerro(str(_y,4)),4)
	else
		return right(lzerro(str(_y,4)),4)+lzerro(str(_m,2))
	endif
case parameters()=3
	if _m=0
		return right(lzerro(str(_y,4)),4)+_pr
	else
		return right(lzerro(str(_y,4)),4)+lzerro(str(_m,2))+_pr
    endif
endcase
return ''

*
* Function LZERRO
* ���������: ����� �����, ������
Function LZerro
parameters _Lz, _Len
_par = parameters()
do case
case _par = 0
	return '0'
case  _par = 1
	_Len = 10
endcase

if type('_Lz')='N'
	_Lz = str(_Lz, _Len)
endif
if type('_Lz')<>'C'
	return '0'
endif

LOCAL _l
_l=len(_Lz)

if len(alltrim(_Lz))<_l
        _Lz=replicate('0',_l-len(alltrim(_Lz)))+alltrim(_Lz)
endif
return (_Lz)


*
* Function GOREC()
* ������������� ��������� �� ����������� ������
* ���������: <����� ������>,[<���������>]
*
Function GoRec
parameters _rec, _als

TRY 
	if parameters()=1
		_als=alias()
	endif
	if _rec>reccount(_als)
		if !eof(_als)
			go bottom in (_als)
			skip in (_als)
		endif
	else
		go _rec in (_als)
	ENDIF
CATCH
	go bottom in (_als)
ENDTRY
return

*
* Function ADDINPATH()
*
Function AddInPath
parameters _p
if parameters() = 0
	return .f.
endif

local _path
_path = set("PATH")
if !empty(_path) and right(_path, 1) # ";"
	_path = _path + ";"
endif
if right(_p, 1) # ";"
	_p = _p + ";"
endif

if atcc(_p, _path) = 0
	_path = _path + _p
endif
set path to (_path)
return .t.

*
* �������� � ������ � ���� DBF ������� ������� ��������
* ��� DOS - 866 (0x65, 101)
*
Function MakeCpRus(_f)
private _han
_han = fopen(_f, 2)
if (_han<0)
	return .f.
endif
fseek(_han, 29,0)
_s = fread(_han,1)
if (_s != "e")
	fseek(_han, 29,0)
	fwrite(_han, "e")
endif
fclose(_han)
return .t.

* -- ���������� �������� ������
Function MonthToC
parameters _m
if pcount() = 0
	_m = month(date())
endif
local _name
do case
case _m = 1
	_name = "������"
case _m = 2
	_name = "�������"
case _m = 3
	_name = "����"
case _m = 4
	_name = "������"
case _m = 5
	_name = "���"
case _m = 6
	_name = "����"
case _m = 7
	_name = "����"
case _m = 8
	_name = "������"
case _m = 9
	_name = "��������"
case _m = 10
	_name = "�������"
case _m = 11
	_name = "������"
case _m = 12
	_name = "�������"
endcase
return _name


* FUNCTION PROPIS_CIEL 
* ����� ������� PROPIS_CIEL(152.35) = 152 ���. 35 ���.
* ���������� ������ ��������
*
Function Propis_Ciel
lparameters lnZn
if type("lnZn") # "N"
	return ""
endif
local lnLeft, lnRight
lnLeft = int(lnZn)
lnRight = int((lnZn - lnLeft) * 100)
return alltrim(str(lnLeft)) + " ���. " + transform(lnRight, "@RL 99 ���.")
endfunc

* PROPIS FUNCTION
* ����� ��������
*
function PROPIS
parameters CIF_
private _0,_1,_2
if CIF_>999999999999.99.or.CIF_<0
        return''
endif
_1=str(CIF_,15,2)
_0=0
_2=SIM(substr(_1,1,3))
do case
case _0=1
        _2=_2+'���� �������� '
case _0=2
        _2=_2+'��� ��������� '
case _0=3
        _2=_2+'��������� '
case _0=4
        _2=_2+'���������� '
endcase
_2=_2+SIM(substr(_1,4,3))
do case
case _0=1
        _2=_2+'���� ������� '
case _0=2
        _2=_2+'��� �������� '
case _0=3
        _2=_2+'�������� '
case _0=4
        _2=_2+'��������� '
endcase
_2=_2+SIM(substr(_1,7,3))
do case
case _0=1
        _2=_2+'���� ������ '
case _0=2
        _2=_2+'��� ������ '
case _0=3
        _2=_2+'������ '
case _0=4
        _2=_2+'����� '
endcase
_2=_2+SIM(substr(_1,10,3))
do case
case _0=0.and.empty(_2)
        _2='���� ������ '
case _0=1
        _2=_2+'���� ����� '
case _0=2
        _2=_2+'��� ����� '
case _0=3
        _2=_2+"����� "
case _0=4
        _2=_2+"������ "
otherwise
        _2=_2+"������ "
endcase
_0 = substr(_1,15,1)
do case
case _0 = "1"
        _0 = " �������"
case _0 >= "2" and _0 <= "4"
        _0 = " �������"
otherwise
        _0 = ' ������'
endcase
ok_ = 100 * (CIF_ - int(CIF_))
local _ret
_ret = _2 + substr(_1,14,2) + _0
if ok_ = 0
	_ret = substr(_ret, 1, rat(" ", _ret, 2))
endif
_ret = lower(_ret)
_ret = stuff(_ret, 1, 1,upper(substr(_ret, 1, 1)))
return _ret


function SIM
parameters SIM_
private CIF,SIM
CIF=val(substr(SIM_,1,1))
SIM=''
do case
case CIF=1
        SIM='��� '
case CIF=2
        SIM='������ '
case CIF=3
        SIM='������ '
case CIF=4
        SIM='��������� '
case CIF=5
        SIM='������� '
case CIF=6
        SIM='�������� '
case CIF=7
        SIM='������� '
case CIF=8
        SIM='��������� '
case CIF=9
        SIM='��������� '
endcase
CIF=val(substr(SIM_,2,1))
do case
case CIF=2
        SIM=SIM+'�������� '
case CIF=3
        SIM=SIM+'�������� '
case CIF=4
        SIM=SIM+'����� '
case CIF=5
        SIM=SIM+'��������� '
case CIF=6
        SIM=SIM+'���������� '
case CIF=7
        SIM=SIM+'��������� '
case CIF=8
        SIM=SIM+'����������� '
case CIF=9
        SIM=SIM+'��������� '
endcase
_0=4
if CIF=1
        CIF=val(substr(SIM_,3,1))
        do case
        case CIF=0
                SIM=SIM+'������ '
        case CIF=1
                SIM=SIM+'����������� '
        case CIF=2
                SIM=SIM+'���������� '
        case CIF=3
                SIM=SIM+'���������� '
        case CIF=4
                SIM=SIM+'������������ '
        case CIF=5
                SIM=SIM+'���������� '
        case CIF=6
                SIM=SIM+'����������� '
        case CIF=7
                SIM=SIM+'���������� '
        case CIF=8
                SIM=SIM+'������������ '
        case CIF=9
                SIM=SIM+'������������ '
        endcase
else
        CIF=val(substr(SIM_,3,1))
        do case
        case CIF=0.and.empty(SIM)
                _0=0
        case CIF=1
                _0=1
        case CIF=2
                _0=2
        case CIF=3
                _0=3
                SIM=SIM+'��� '
        case CIF=4
                _0=3
                SIM=SIM+'������ '
        case CIF=5
                SIM=SIM+'���� '
        case CIF=6
                SIM=SIM+'����� '
        case CIF=7
                SIM=SIM+'���� '
        case CIF=8
                SIM=SIM+'������ '
        case CIF=9
                SIM=SIM+'������ '
        endcase
endif
return SIM

* ������ ��������� ���� � ��������
* WSave(oFormRef)
* oFormRef			����� ���� (Thisform)
* ������: WSave(thisform)
*
Function WSave(oFormRef as Form) as Boolean
local _sel, cFile
_sel = select()
cFile = SYS(2023)+"\WPOS.DBF"
SELECT 0
IF !FILE(cFile)
	CREATE TABLE (cFile) FREE (WId C(50), Top I, Left I, Width I, Height I)
	INDEX ON Wid TAG WID
	USE IN WPOS
ENDIF
USE (cFile) AGAIN ALIAS WPOS
if indexseek(alltrim(m.oFormRef.name), .t., "WPOS", "WID")
	replace wpos.Top with m.oFormRef.Top, wpos.Left with m.oFormRef.Left,;
		wpos.Width with m.oFormRef.Width, wpos.Height with m.oFormRef.Height IN WPOS
ELSE
	insert into WPOS (WId, Top, Left, Width, Height) values (m.oFormRef.name, m.oFormRef.Top,;
		m.oFormRef.Left, m.oFormRef.Width, m.oFormRef.Height)
endif
USE IN WPOS
select(_sel)
return (.t.)

* ������ ��������� ���� � ��������, ��� ������ ����������� ������� WSave(iUserID, oFormRef)
* WGet(oFormRef)
* oFormRef		����� ���� (This)
* ������: WGet(thisform)
*
Function WGet(oFormRef as Form) as Boolean
local _sel, cFile, oExp as Exception, lRet as Boolean
_sel = select()
cFile = SYS(2023)+"\WPOS.DBF"

IF !FILE(cFile)
	Wsave(oFormRef)
ENDIF
oExp = NULL
lRet = .T.
TRY
	SELECT 0
	USE (cFile) AGAIN ALIAS WPOS

	IF INDEXSEEK(alltrim(m.oFormRef.name), .t., "WPOS", "WID")
		m.oFormRef.Top = wpos.Top
		m.oFormRef.Left = wpos.Left
		IF m.oFormRef.BorderStyle = 3
			m.oFormRef.Width = wpos.Width
			m.oFormRef.Height = wpos.Height
		ENDIF
	ENDIF
	USE IN WPOS
	select(_sel)

CATCH TO oExp
	lRet = .F.
	m.oFormRef.Top = 1
	m.oFormRef.Left = 1
	
FINALLY
	IF USED("WPOS")
		USE IN WPOS
	ENDIF
ENDTRY
return (lRet)

* ������� �������� Period_Dnei()
* ���������� ����: DD.MM.YYYY
FUNCTION Period_Dnei_Dt
LPARAMETERS lcStr as Character
PRIVATE _Dd, _Mm, _Yy
_Dd = 1
DO CASE
CASE ATC("������",lcStr)>0 OR ATC("������",lcStr)>0 OR ATC("�����",lcStr)>0
	_Mm = 1
CASE ATC("�������",lcStr)>0 OR ATC("�������",lcStr)>0 OR ATC("������",lcStr)>0
	_Mm = 2
CASE ATC("����",lcStr)>0 OR ATC("�����",lcStr)>0
	_Mm = 3
CASE ATC("������",lcStr)>0 OR ATC("������",lcStr)>0 OR ATC("�����",lcStr)>0
	_Mm = 4
CASE ATC("���",lcStr)>0 OR ATC("���",lcStr)>0
	_Mm = 5
CASE ATC("����",lcStr)>0 OR ATC("����",lcStr)>0
	_Mm = 6
CASE ATC("����",lcStr)>0 OR ATC("����",lcStr)>0
	_Mm = 7
CASE ATC("������",lcStr)>0 OR ATC("�������",lcStr)>0
	_Mm = 8
CASE ATC("��������",lcStr)>0 OR ATC("��������",lcStr)>0 OR ATC("������",lcStr)>0
	_Mm = 9
CASE ATC("�������",lcStr)>0 OR ATC("�������",lcStr)>0 OR ATC("�����",lcStr)>0
	_Mm = 10
CASE ATC("������",lcStr)>0 OR ATC("������",lcStr)>0 OR ATC("�����",lcStr)>0
	_Mm = 11
CASE ATC("�������",lcStr)>0 OR ATC("�������",lcStr)>0 OR ATC("������",lcStr)>0
	_Mm = 12
ENDCASE
_Yy = VAL(SUBSTR(lcStr,AT(SPACE(1),lcStr,5)+1,4))
RETURN (DATE(_Yy,_Mm, 1))
ENDFUNC

* ����������� ������� ���� � ������, ��������: � 01 �� 30 ������ 2013 ����
* Period_Dnei(ldDt)
* ldDt - ����, �� ���� ������� ������ ����� � ���
FUNCTION Period_Dnei
LPARAMETERS _dp
LOCAL _m, _y, _lday
_m = MONTH(_dp)
_y = YEAR(_dp)
_lday = LDay(_dp)

RETURN "� 01 �� "+TRANSFORM(_lday,"@L 99")+" "+RusMonth_EX(_m,2)+" "+TRANSFORM(_y, "@L 9999")+" ����"
ENDFUNC

* ����������� ���-�� ���� � ������
* LDay(ldDt)
* ldDt - ����
*
FUNCTION LDay
lparameters _dp
private _ret,_m
IF TYPE('_dp')<>"D"
	_dp = DATE()
ENDIF
_m = month(_dp)
do case
case _m=1
        _ret=31
case _m=2
        _ret=iif(vis(_dp),29,28)
case _m=3
        _ret=31
case _m=4
        _ret=30
case _m=5
        _ret=31
case _m=6
        _ret=30
case _m=7
        _ret=31
case _m=8
        _ret=31
case _m=9
        _ret=30
case _m=10
        _ret=31
case _m=11
        _ret=30
case _m=12
        _ret=31
endcase
return _ret

* ����������� ���������� ��� ��� ���
* Vis(dDt)
* dDt ����
*
function Vis
parameters _dp
private _y,_ret
_y=year(_dp)
_ret=.F.
if mod(_y,4)=0
	_ret=.T.
	if mod(_y,100)=0
		_ret=.F.
		if mod(_y,400)=0
			_ret=.T.
		endif
	endif
endif
return _ret

*
* ��������� ������ ��� �������
* ���������:
* 1 - ������
* 2 - ������ ������ ������ getfont()
*
Function SetFont
lparameters _O, _Font
if pcount() < 2
	return .f.
endif
if isnull(_O)
	return .f.
endif
if empty(_Font)
	_Font = getfont()
	if empty(_Font)
		return .f.
	endif
endif
_FontName = GetFontName(_Font)
_FontSize = GetFontSize(_Font)
_FontStyle = GetFontStyle(_Font)
with _O
	.FontName = _FontName
	.FontSize = val(_FontSize)
	do case
	case _FontStyle == "N"
		.FontBold = .f.
		.FontItalic = .f.
	case _FontStyle == "I"
		.FontItalic = .t.
	case _FontStyle == "B"
		.FontBold = .t.
	case _FontStyle == "BI"
		.FontBold = .t.
		.FontItalic = .t.
	endcase
endwith
return .t.

Function GetFontName
lparameters _Font
if empty(_Font)
	return "Arial"
endif
local _c1, _c2, _l
_c1 = at(",", _Font, 1)
_c2 = at(",", _Font, 2)
_l = len(_Font)
return substr(_Font, 1, _c1-1)

Function GetFontSize
lparameters _Font
if empty(_Font)
	return "8"
endif
local _c1, _c2, _l
_c1 = at(",", _Font, 1)
_c2 = at(",", _Font, 2)
_l = len(_Font)
return substr(_Font, _c1+1, _l-_c1-2)

Function GetFontStyle
lparameters _Font
if empty(_Font)
	return "N"
endif
local _c1, _c2, _l
_c1 = at(",", _Font, 1)
_c2 = at(",", _Font, 2)
_l = len(_Font)
return substr(_Font, _c2+1)


*
* ������ ���������� �� ����� access.ass
* ���������
* 1 - ������, �������� "[Test]"
* 2 - ��������, �������� "FontName"
*
Function GetAss
lparameters _Section, _VarName
if !file(_PathHome+"\ACCESS.DBF")
	return ""
endif
local _sel, _ret, i, j
_ret = ""
_sel = select()
if used("ASS")
	select ASS
else
	select 0
	use (_PathHome+"\ACCESS.DBF") again order 1 alias ASS
endif
if indexseek(_UserPass, .t., "ASS")
	if alines(_myAss, ass.ass) > 0
		for i=1 to alen(_myAss)
			if _myAss[i] == _Section
				i = i + 1
				for j=i to alen(_myAss)
					* �������� ������ ������ ?
					if at("[", _myAss[j]) # 0 and at("]", _myAss[j]) # 0
						exit
					endif
					if at(_VarName, _myAss[j]) > 0
						_ret = substr(_myAss[j], at("=", _myAss[j]) + 1)
					endif
				next
				i = j
			endif
		next
	endif
endif
use in ASS
select(_sel)
return _ret


*
* ������ ���������� � ����� access.ass
* ���������
* 1 - ������, �������� "[Test]"
* 2 - ��������, �������� "FontName"
* 3 - ��������
*
Function SetAss
lparameters _Section, _VarName, _Value
if !file(_PathHome+"\ACCESS.DBF")
	return ""
endif
local _myAss, _sel, _ret, i, j, _temp, _found
dime _myAss[1]
_ret = ""
_found = .f.
_sel = select()
if used("ASS")
	select ASS
else
	select 0
	use (_PathHome+"\ACCESS.DBF") again order 1 alias ASS
endif

if indexseek(_UserPass, .t., "ASS")
	if alines(_myAss, ass.ass) > 0
		for i=1 to alen(_myAss)
			dime _temp[i]
			_temp[i] = _myAss[i]
			* ����� ������� ������
			if _myAss[i] == _Section
				i = i + 1
				for j=i to alen(_myAss)
					* �������� ������ ������ ?
					if at("[", _myAss[j]) # 0 and at("]", _myAss[j]) # 0
						dime _temp[j+1]
						_temp[alen(_temp)] = _varName + " = " + _Value
						_found = .t.
						exit
					endif
					dime _temp[j]
					_temp[j] = _myAss[j]
					if at(_VarName, _myAss[j]) > 0
						_temp[j] = _varName + " = " + _Value
						_found = .t.
					endif
				next
				i = j
			endif			
		next
		if !_found
			dime _temp[alen(_temp)+2]
			_temp[alen(_temp)] = _Varname + " = " + _Value
			_temp[alen(_temp)-1] = _Section
		endif
	else
		dime _temp[2]
		_temp[1] = _Section
		_temp[2] = _Varname + " = " + _Value
	endif
	replace ass.ass with ""
	for i=1 to alen(_temp)
		replace ass.ass with _temp[i] + chr(13) additive
	next
endif
use in ASS
select(_sel)
return _ret


*!*	*** ������� GUID() - ���������� 8 �������� ���������� �������� 
*!*	*** GUID - ������� �� ������� (� �����. �� ���� 0,1 ��), ������ � ���� ��
*!*	*** ����� ��������: DatFromGuid() - �������������� ���� �� GUID
*!*	*** ���������� ���������� WSUBD.FLL (SET LIBRARY TO ...)
*!*	*** ��������� ��������������, (�) ��.������, 1998


*!*	****************************************************************
*!*	*** ������������� ���������� GUID - �������� � ������ ����������
*!*	Function IniGenGuid
*!*	Public hid2_guid, UniString_, nd96s, nd96d && �������� ��� �������������� GUID
*!*	 nd96s=0
*!*	 nd96d=0
*!*	 *** �������� ������� �������� ��� ����������� ����
*!*	 UniString_=GetUniString() && �������� ��� �������� ����������
*!*	 *** ������������� 
*!*	 cBuf=Sys(2015) && ��������� � ���������� ��� ��� ������
*!*	  ** GetVolInfo - �������� ���������� � HARD (hdd)
*!*	  ** GetICcrc16 - �������� �����. ����� ������
*!*	  ** NumToB     - ��������� ����� � ������
*!*	 ** Hid2_Guid - 2 ����� �����. ����� �� � ������ �������
*!*	 Hid2_Guid=PADR(NumtoB(GetICrc16(GetVolInfo()+cBuf)),2,'*')&& low ������� - � ����� wsubd.fll
*!*	 nd96s=Tic96() && ��������� �������� �����
*!*	RETURN

*!*	*************************************************************************
*!*	* �������� ������� - �������� ���������� ������������� GUID = Hard+Time
*!*	Function GUID
*!*	* �� 2010 ����
*!*	* 6b ��������� ������� + 2b hard-������������+������
*!*	Return PADR(NumtoB(Tic96n())+Hid2_guid,8,'!')


*!*	************************************
*!*	** ������������ ���������� ����� �� GUID (���� + 2������� ������� + 2 ������� ������� ���)
*!*	Function NumFromGuid(cB)
*!*	RETURN BTONum(SUBSTR(cb,1,6))


*!*	************************************
*!*	** ������������ ���� �� Guid
*!*	Function DatFromGuid(cB)
*!*	Local nTic, dDate
*!*	nTic=BTONum(SUBSTR(cb,1,6))
*!*	nTic=nTic/100 && ������ 2 ���.������� ��������
*!*	dDate=INT(nTic/8640000) && ���
*!*	dDate=CTOD(SYS(10,dDate+2450800))
*!*	RETURN dDate


*!*	*********************************************************
*!*	*** ���������� ������� �������� ��� ����������� ����
*!*	Function GetUniString
*!*	Local cUniS,ni
*!*	    && cUnis - ��� ����� ������ ��� ����������� ����� � ���������� LEN(cInis)
*!*	    && �������� �������: '"[]&   (39,34,91,93)
*!*	    *** �������� ������ ������
*!*	            cUnis=chr(35)+chr(36)+chr(37) && +chr(38) - "&"
*!*	   For ni=40 to 90
*!*	       cUniS=cUniS+chr(ni)
*!*	   EndFor
*!*	            cUniS=cUniS+chr(92)
*!*	   For ni=94 to 126
*!*	       cUniS=cUniS+chr(ni)
*!*	   EndFor
*!*	   For ni=192 to 255
*!*	       cUniS=cUniS+chr(ni)
*!*	   EndFor
*!*	Return cUniS


*!*	*********************************************
*!*	* �������� ����� ����� (10 �� - ������ �� ����� NT) ��������� � 1998 ���� 
*!*	* 6 ���� - ������� �� 20 ���
*!*	* 2451545 - ����� ����, ��������� � 01-01-2000
*!*	Function Tic96
*!*	return int(val(sys(11,date()-2451545))*8640000+seconds()*100)


*!*	**********************************************************
*!*	* �������� ���������� ����� - ���� Tic96 + �������-����������
*!*	* � �������� ������ ���� Tic96 ����� ��������� ��������� ���
*!*	* ������� ������ ������������� increment, old - in nd96
*!*	Function Tic96N
*!*	Local nNew,nNewS
*!*	 nNewS=Tic96()
*!*	  IF nNewS==nd96s && ��� ��� ������ ���
*!*	    IF nd96d==99 && ��� ��� ������ �������� - ����� ����� ���
*!*	       DO WHILE nNewS==nd96s
*!*	         nNewS=Tic96()
*!*	       ENDDO
*!*	      nd96d=0
*!*	    ELSE
*!*	      nd96d=nd96d+1 && ����� ���� - ����������� �������
*!*	    ENDIF  
*!*	  ELSE && ����� ���
*!*	    nd96d=0  
*!*	  ENDIF
*!*	       nNew=nNewS*100+nd96d  && ������� ����� Id
*!*	       nd96s=nNewS
*!*	return (nNew)


*!*	************************************************************
*!*	* �������������� ����� nNum � ������ ��������
*!*	* ��� ���������� �������� ����� � ������ ��������
*!*	* ������ ������������ ������ ��������� � ���������� UniString_
*!*	* (c) Dm.Bayanov 1997 , 1998
*!*	* (�-� VFP BinToC() �������� ��� ����� �������� > 9)
*!*	* �������� �������������� - BtoNum
*!*	*** ������ ������ ������ ������������
*!*	*UniString_="01" - 2
*!*	*UniString_="0123456789" - 10
*!*	*UniString_="0123456789ABCDEF" - 16
*!*	Function NumtoB(nNum)
*!*	Local cBuf,ni,nOst,nTale,nOsn,nVes, nNum, nLen, cUniStr
*!*	cBuf=""
*!*	nOst=0
*!*	nTale=0
*!*	nLen=1
*!*	cUniStr=UniString_
*!*	*** ��������� ��������� ����� - ����� ������� ������������ ������ UniString
*!*	nOsn=Len(cUniStr)&&-1 && �������� ���������� �������� ��� ������
*!*	*** �������� ���������� �������� � �������� ����� (ex.999>123)
*!*	DO WHILE .T.
*!*	    nOst=nOst+(nOsn-1)*nOsn**(nLen-1)
*!*	 IF nOst >= nNum
*!*	   EXIT
*!*	 ENDIF
*!*	  nLen=nLen+1
*!*	ENDDO
*!*	   FOR ni=1 TO nLen
*!*	     nOst=nNum-nTale
*!*	      * ��� ni ������� - ������ ������ ����� ��������� ��������
*!*	      *  �� �������� �� �������� �������� UniString
*!*	     nVes=INT(nOst/nOsn**(nLen-ni))
*!*	      * ������� ������� �� ������� �������� �������
*!*	     nTale=nTale+nVes*nOsn**(nLen-ni)
*!*	         *  �� ���� ������� ��������� ������ �� ������ ���������� ������ UniStr
*!*	         * nVes+1 ��� ��������� ����� �� nVes==0
*!*	        cBuf=cBuf+SUBSTR(cUniStr,nVes+1,1)
*!*	   NEXT
*!*	RETURN cBuf


*!*	*********************************************
*!*	** �������� �������������� � ���������� �����
*!*	Function BtoNum(cB)
*!*	Local ni,nj,nOst,nk,nOsn,cCim,cUniStr
*!*	nOst=0
*!*	cUniStr=UniString_
*!*	nOsn=LEN(cUniStr)
*!*	FOR ni=1 TO LEN(cB)
*!*	  cCim=SUBSTR(cb,LEN(cb)-ni+1,1)
*!*	  nk=AT(cCim,cUniStr)-1
*!*	  nOst=nOst+nk*(nOsn**(ni-1))
*!*	NEXT
*!*	RETURN nOst


*
* ������ � ���� INI
*
* ���������:
* tcFileName		- ���� � ����������� .INI, ���� ������ ���� ��� ����, �� ������� � ������� %OS%
* tcSection			- ������
* tcEntry				- ��������
* tcValue				- �������� ���������
* ����������:
* .t. - OK ����� .f.
*
FUNCTION WriteFileIni(tcFileName,tcSection,tcEntry,tcValue)
DECLARE INTEGER WritePrivateProfileString ;	
	IN WIN32API ;
	STRING cSection,STRING cEntry,STRING cEntry,;	
	STRING cFileName
RETURN IIF(WritePrivateProfileString(tcSection,tcEntry,tcValue,tcFileName)=1, .T., .F.)

ENDFUNC

*
* ������ �� ����� INI
*
* ��������:
* tcFileName			- ���� � ����������� .INI, ���� ������ ���� ��� ����, �� ������ �� �������� %OS%
* tcSection				- ������
* tcEntry				- ��������
* ����������:
* �������� ��������� tcEntry ��� .NULL.
*
FUNCTION ReadFileIni(tcFileName,tcSection,tcEntry)
LOCAL lcIniValue, lnResult, lnBufferSize
DECLARE INTEGER GetPrivateProfileString ;    
	IN WIN32API ;    
	STRING cSection,;
    STRING cEntry,;    
    STRING cDefault,;    
    STRING @cRetVal,;    
    INTEGER nSize,;
    STRING cFileName

lnBufferSize = 255
lcIniValue = space(lnBufferSize)

lnResult=GetPrivateProfileString(tcSection,tcEntry,"*NULL*",@lcIniValue,lnBufferSize,tcFileName)
lcIniValue=SUBSTR(lcIniValue,1,lnResult)

IF lcIniValue=="*NULL*"
	tcRet = .NULL.
ELSE
	tcRet = lcIniValue
ENDIF 

RETURN tcRet
ENDFUNC

*
* ������������ �� ������� � �������������� ������
*
FUNCTION ActiKeybrdRus
LOCAL lpr
DECLARE SHORT GetKeyboardLayoutName IN user32.dll STRING @lpR
DECLARE SHORT ActivateKeyboardLayout IN user32.dll INTEGER HKL , INTEGER flags
lpr='           '
GetKeyboardLayoutName(@lpr)
if not '419' $lpr
	ActivateKeyboardLayout(1,0)
endif 
CLEAR DLLS "GetKeyboardLayoutName", "ActivateKeyboardLayout"
RETURN
ENDFUNC

*
* ������������ �� US ��������� ����������
*
FUNCTION ActiKeybrdUS
LOCAL lpr
DECLARE SHORT GetKeyboardLayoutName IN user32.dll STRING @lpR
DECLARE SHORT ActivateKeyboardLayout IN user32.dll INTEGER HKL , INTEGER flags
lpr='           '
GetKeyboardLayoutName(@lpr)
if not '409' $lpr
	ActivateKeyboardLayout(1,0)
endif 
CLEAR DLLS "GetKeyboardLayoutName", "ActivateKeyboardLayout"
RETURN
ENDFUNC

*
* ������� ��� ��������� ����������
* 409 - ����
* 419 - ���.
*
FUNCTION GetKeybrdLayout
LOCAL lpr
DECLARE SHORT GetKeyboardLayoutName IN user32.dll STRING @lpR
lpr='           '
GetKeyboardLayoutName(@lpr)
CLEAR DLLS "GetKeyboardLayoutName"
RETURN (VAL(lpr))
ENDFUNC


*
* ������ ������ �� ������� ������ ������� ��������� � �������� ��
*
FUNCTION GetTableInDbc
LOCAL cF
cF = SYS(2023)+"\"+SYS(3)+".LIST DATABASE"
list database to FILE (cF) noconsole

CREATE CURSOR TABLEMAP (DBC C(50), TABLE C(50), PATH C(50), Coment C (50))

_han = fopen(cF)
if _han < 0
	RETURN ("")
endif
do while !feof(_han)
	store "" to m.dbc, m.table, m.path, m.coment
	m.dbc = justfname(dbc())
	_s = fgets(_han)
	if at("Table", _s) = 1
		m.table = alltrim(substr(_s, len("Table")+1))
		do while !feof(_han)
			_s = fgets(_han)
			if at("*Path", _s) # 0
				m.path = alltrim(substr(_s, at("*Path", _s)+len("*Path")+1))
			ENDIF
			IF AT("*Comment", _s) # 0
				m.coment = alltrim(substr(_s, at("*Comment", _s)+len("*Comment")+1))
				exit
			ENDIF
		enddo
		locate for tablemap.table = m.table and tablemap.path = m.path
		if !found()
			insert into TableMap from memvar
		endif
	endif
enddo
=fclose(_han)
return("TABLEMAP")


*! ����������� �������� �����
*!
FUNCTION NetMap
DECLARE DOUBLE WNetAddConnection IN MPR.DLL ; 
STRING @ lpszRemoteName,; 
STRING @ lpszPassword,; 
STRING @ lpszLocalName 

*share name of the mapping 
RemoteName="\\ShareName" 
*logged user password, if its the same password set to .null. 
CurrentUserPassword= "ThePassword" 
*Map to this local drive letter 
LocalDriveLetter="K:" 

=WNetAddConnection(@RemoteName,@CurrentUserPassword,@LocalDriveLetter) 
RETURN .t.
ENDFUNC


* CALL     RIGHT_P (<expC1>)
* <STR_1> ������ ��������
* ?RIGHT_P('�������� ����� �����������')
* �������� �.�.
*
FUNCTION RIGHT_P(_str)
LOCAL i,_a,j
_str=alltrim(_str)
i=at(' ',_str)
if empty(_str) or i=0
        return _str
endif
_a=substr(_str,1,i)
for j=i to len(_str)
        do while empty(substr(_str,j,1))
                j=j+1
        enddo
        _a=_a+substr(_str,j,1)+'.'
        do while !empty(substr(_str,j,1))
                j=j+1
        enddo
endfor
return _a

*
* �������� ���������� ������� ���������
*
* IF FirstInstance("Narkota") == .T.
* 	DO FORM SYS(5)+SYS(2003)+'\Narkota'
* 	READ EVENTS
* ELSE
* 	MessageBox("�� ��� ����������� ���� ���! ������ ����������.")
* ENDIF
*
FUNCTION FirstInstance() 
LPARAMETERS lcStarBase 

#DEFINE SW_RESTORE 9 
#DEFINE ERROR_ALREADY_EXISTS 183 
#DEFINE GW_HWNDNEXT 2 
#DEFINE GW_CHILD 5 

DECLARE INTEGER CreateMutex IN win32api INTEGER, INTEGER, STRING @
DECLARE INTEGER CloseHandle IN win32api INTEGER
DECLARE INTEGER GetLastError IN win32api
DECLARE INTEGER SetProp IN win32api INTEGER, STRING @, INTEGER
DECLARE INTEGER GetProp IN win32api INTEGER, STRING @
DECLARE INTEGER RemoveProp IN win32api INTEGER, STRING @
DECLARE INTEGER IsIconic IN win32api INTEGER
DECLARE INTEGER SetForegroundWindow IN USER32 INTEGER
DECLARE INTEGER GetWindow IN USER32 INTEGER, INTEGER
DECLARE INTEGER ShowWindow IN WIN32API INTEGER, INTEGER
DECLARE INTEGER GetDesktopWindow IN WIN32API

LOCAL cExeFlag											&& Name of the MUTEX 
LOCAL nExeHwnd											&& MUTEX handle 
LOCAL hWnd												&& Window handle 
LOCAL lRetVal											&& Return value of this function

IF EMPTY(lcStarBase)
	cExeFlag = "Narkota"+CHR(0)
ELSE
	cExeFlag = lcStarBase+CHR(0)
ENDIF
nExeHwnd = CreateMutex(0,1,@cExeFlag)

IF GetLastError() = ERROR_ALREADY_EXISTS
	HWND = GetWindow(GetDesktopWindow(), GW_CHILD)
	DO WHILE HWND > 0
		IF GetProp(HWND, @cExeFlag) = 1
			IF IsIconic(HWND) > 0
				ShowWindow(HWND,SW_RESTORE)
			ENDIF
			SetForegroundWindow(HWND)
			EXIT
		ENDIF
		HWND = GetWindow(HWND,GW_HWNDNEXT)
	ENDDO
	CloseHandle(nExeHwnd)
	lRetVal = .F.
ELSE
	SetProp(1, @cExeFlag, 1)
	lRetVal = .T.
ENDIF
CLEAR DLLS 'CreateMutex', 'CloseHandle', 'GetLastError', 'SetProp', 'GetProp', 'RemoveProp', 'IsIconic', 'SetForegroundWindow', 'GetWindow', 'GetDesktopWindow'
RETURN lRetVal
ENDFUNC

FUNCTION quest(cMSG)
IF PCOUNT()=0
	cMSG = REPLICATE('?',50)
	FOR nI=1 TO 9
		cMSG = cMSG+CHR(13)+REPLICATE('?',50)
	NEXT
ENDIF
LOCAL oForm as Form
oForm = CREATEOBJECT('oMSGFORM',cMSG)
oForm.show(1)
ENDFUNC

DEFINE CLASS oMSGFORM as Form
	caption = '���������'
	AutoCenter = .T.
	Width = 200
	Height = 100
	BorderStyle = 2
	DataSession = 1
	ShowWindow = 1
	ShowTips = .T.
	MaxButton = .F.
	MinButton = .F.
	
	ADD OBJECT lblMSG as label WITH top=2,left=2,autosize=.t.,caption='',visible=.t.
	
ENDDEFINE

*********************************************************
* ������� �������������� �������, ���, ��������
* � ����� ������� ������
*
* ���������:
* mFamily - �������
* mName - ���
* mSoName - ��������
* Pol - ��� (���������� �������� "�", "�")
* Padeg - ����� (���������� �������� 1..6)
* && 1-������������
* && 2-�����������
* && 3-���������
* && 4-�����������
* && 5-������������
* && 6-����������
*********************************************************

Function FIO
Parameter mFamily, mName, mSoName, Pol, Padeg

mFamily=AllTrim(mFamily)
mName=AllTrim(mName)
mSoName=AllTrim(mSoName)

If Between(Padeg, 1, 6)
	Pol=Lower(Pol)
	Do Case
	Case Pol="�"
		mFio=Fam_m(mFamily, padeg) && ����������� �������
		mFio=mFio + " " + Name_m(mName, padeg) && ����������� ���
		mFio=mFio + " " + SoName_m(mSoName, padeg) && ����������� ��������
	Case Pol="�"
		mFio=Fam_w(mFamily, padeg) && ����������� �������
		mFio=mFio + " " + Name_w(mName, padeg) && ����������� ���
		mFio=mFio + " " + SoName_w(mSoName, padeg) && ����������� ��������
	OtherWise && ���� �� ������ ���, ������ ��������� �� ������
		mFio="H������ ������ ��� !"
	EndCase
ELSE
	mFio="H����������� �������� ������ !"
EndIf
Return mFio




********************************************
* �������������� ������� �������
********************************************
Function Fam_m
Parameter Fam, Pad && ��������� ������� � ������. ������
&& � ��������� �����
&& 1-������������
&& 2-�����������
&& 3-���������
&& 4-�����������
&& 5-������������
&& 6-����������

Declare Oki(6) && ���������� ������ ���������
Store "" to Oki

Fam_m=""
Fam=Proper(Fam)
Len=Len(Fam) && ���������� ���������� ���� � �������

Ok=Substr(Fam, Len-2, 3) && ����� ��������� 3 ����� �������
If Ok="���" .OR. Ok="���"; && ���� ��������� �����

Fam=Substr(Fam, 1, len-2) && �� ��������� �������
Do Case
Case Ok="���"
	Oki[1]="��"
	Oki[2]="���"
	Oki[3]="���"
	Oki[4]="���"
	Oki[5]="��"
	Oki[6]="��"
Case Ok="���"
	Oki[1]="��"
	Oki[2]="���"
	Oki[3]="���"
	Oki[4]="���"
	Oki[5]="��"
	Oki[6]="��"
EndCase
Fam_m=Fam+Oki[Pad] && ��������� � ������� ���������
Return Fam_m && ��������� ���������
EndIf

***** ���������� ������ ���������
Ok=Substr(Fam, Len-1, 2) && ����� ��������� ��� �����
If Ok="��"
	Fam=Substr(Fam, 1, Len-1) && ��������� �������
	Oki[1]="�"
	Oki[2]="�"
	Oki[3]="�"
	Oki[4]="�"
	Oki[5]="��"
	Oki[6]="�"
	Fam_m=Fam+Oki[Pad] && ��������� � ������� ���������
	Return Fam_m && ��������� ���������
EndIf

Do Case
Case Ok="��" .OR. Ok="��" .OR. Ok="��" .OR. Ok="��" .OR. Ok="��"
	Oki[1]=""
	Oki[2]="�"
	Oki[3]="�"
	Oki[4]="�"
	Oki[5]="��"
	Oki[6]="�"
Case Ok="��" .OR. Ok="��" .OR. Ok="��" .OR. Ok="��" .OR. Ok="��" .OR. Ok="��"
	Oki[1]=""
	Oki[2]="�"
	Oki[3]="�"
	Oki[4]="�"
	Oki[5]="��"
	Oki[6]="�"
Case Ok="��"
	If Pad>1
		Fam=Substr(Fam, 1 , len-2)+"�"
	EndIf
	Oki[1]=""
	Oki[2]="�"
	Oki[3]="�"
	Oki[4]="�"
	Oki[5]="��"
	Oki[6]="�"
Case Ok="��" .OR. Ok="��"
	Fam=Substr(Fam, 1, Len-2)
	Oki[1]=""
	Oki[2]="���"
	Oki[3]="���"
	Oki[4]="���"
	Oki[5]="��"
	Oki[6]="��"
Case Ok="��"
	Oki[1]=""
	Oki[2]="�"
	Oki[3]="�"
	Oki[4]="�"
	Oki[5]="��"
	Oki[6]="�"
EndCase
Fam_m=Fam+Oki[Pad]
Return Fam_m



********************************************
* �������������� ������� �������
********************************************
Function Fam_w

Parameter Fam, Pad && ��������� ������� � ����. ������
&& � ��������� �����
&& 1-������������
&& 2-�����������
&& 3-���������
&& 4-�����������
&& 5-������������
&& 6-����������

Declare Oki(6) && ���������� ������ ���������
Store "" to Oki

Fam_m=""
Fam=Proper(Fam)
Len=Len(Fam) && ���������� ���������� ���� � �������

Ok=Substr(Fam, Len-1, 2) && ����� �������� 3 ����� �������
If Ok="��" .OR. Ok="��"; && ���� ��������� �����

Fam=Substr(Fam, 1, len-2) && �� ��������� �������
Do Case
Case Ok="��"
Oki[1]="��"
Oki[2]="��"
Oki[3]="��"
Oki[4]="��"
Oki[5]="��"
Oki[6]="��"
Case Ok="��"
Oki[1]="��"
Oki[2]="��"
Oki[3]="��"
Oki[4]="��"
Oki[5]="��"
Oki[6]="��"
EndCase
Fam_m=Fam+Oki[Pad] && ��������� � ������� ���������
Return Fam_m && ��������� ���������
EndIf


***** ���������� ������ ���������

Ok=Substr(Fam, Len-2, 3) && ����� ��������� ��� �����

If Ok="���" .OR. Ok="���" .OR. Ok="���" .OR. Ok="���"
Fam=Substr(Fam, 1, Len-1) && ��������� �������
Oki[1]=Ok
Oki[2]="��"
Oki[3]="��"
Oki[4]="�"
Oki[5]="��"
Oki[6]="��"
EndIf

Fam_m=Fam+Oki[Pad]
Return Fam_m



**************************************************
* ������� �������������� �������� �����
*
*
**************************************************

Function Name_m
Parameter Name, Pad && ���������: ��� � ����. ������
&& � ��������� �����
&& 1-������������
&& 2-�����������
&& 3-���������
&& 4-�����������
&& 5-������������
&& 6-����������

Declare Oki(6) && ���������� ������ ���������
Store "" to Oki

Name=Proper(Name)
Len=Len(Name)

Ok=Substr(Name, Len, 1) && ������� ���������
Do Case
Case Ok="�"
Oki[1]=""
Oki[2]="�"
Oki[3]="�"
Oki[4]="�"
Oki[5]="��"
Oki[6]="�"
Name=Substr(Name, 1, Len-1)
OtherWise
Oki[1]=""
Oki[2]="�"
Oki[3]="�"
Oki[4]="�"
Oki[5]="��"
Oki[6]="�"
EndCase
Name_m=Name+Oki[Pad]
Return Name_m



**************************************************
* ������� �������������� �������� �����
*
*
**************************************************

Function Name_w
Parameter Name, Pad && ���������: ��� � ����. ������
&& � ��������� �����
&& 1-������������
&& 2-�����������
&& 3-���������
&& 4-�����������
&& 5-������������
&& 6-����������

Declare Oki(6) && ���������� ������ ���������
Store "" to Oki

Name=Proper(Name)
Len=Len(Name)

Ok=Substr(Name, Len, 1) && ������� ���������
Do Case
Case Ok="�"
Oki[1]=""
Oki[2]="�"
Oki[3]="�"
Oki[4]="�"
Oki[5]="��"
Oki[6]="�"
Name=Substr(Name, 1, Len-1)
Case Ok="�"
Oki[1]=""
Oki[2]="�"
Oki[3]="�"
Oki[4]="�"
Oki[5]="��"
Oki[6]="�"
Name=Substr(Name, 1, Len-1)
EndCase
Name_m=Name+Oki[Pad]
Return Name_m




**************************************************
* ������� �������������� �������� ��������
*
*
**************************************************

Function SoName_m
Parameter SoName, Pad && ���������: ��� � ����. ������
&& � ��������� �����
&& 1-������������
&& 2-�����������
&& 3-���������
&& 4-�����������
&& 5-������������
&& 6-����������

Declare Oki(6) && ���������� ������ ���������
Store "" to Oki

SoName=Proper(SoName)
Len=Len(SoName)

Oki[1]=""
Oki[2]="�"
Oki[3]="�"
Oki[4]="�"
Oki[5]="��"
Oki[6]="�"

SoName_m=SoName+Oki[Pad]
Return SoName_m




**************************************************
* ������� �������������� �������� ��������
*
*
**************************************************

Function SoName_w
Parameter SoName, Pad && ���������: ��� � ����. ������
&& � ��������� �����
&& 1-������������
&& 2-�����������
&& 3-���������
&& 4-�����������
&& 5-������������
&& 6-����������

Declare Oki(6) && ���������� ������ ���������
Store "" to Oki

SoName=Proper(SoName)
Len=Len(SoName)

Ok=Substr(SoName, Len, 1) && ������� ���������
If Ok="�"
Oki[1]="�"
Oki[2]="�"
Oki[3]="�"
Oki[4]="�"
Oki[5]="��"
Oki[6]="�"
SoName=Substr(SoName, 1, Len-1)
EndIf
SoName_w=SoName+Oki[Pad]
Return SoName_w

*
* �������������� ������� ���� � ��������. ������� 1
*
FUNCTION Rus2Translit
LPARAMETERS lpValue  
  CREATE CURSOR tBukv (cKir C(2), cLat C(2))  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","c")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","u")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","k")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","e")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","n")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","g")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","sh")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","sh")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","z")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","h")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","f")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","i")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","v")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","a")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","p")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","r")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","o")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","l")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","d")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","j")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","e")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","ya")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","ch")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","s")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","m")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","i")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","t")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","b")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","yu")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","")  
 *********************************************************  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","Y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","C")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","U")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","K")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","E")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","N")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","G")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","SH")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","SH")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","Z")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","H")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","F")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","I")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","V")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","A")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","P")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","R")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","O")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","L")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","D")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","J")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","E")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","YA")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","CH")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","S")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","M")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","I")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","T")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","Y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","B")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","YU")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("�","")  
 *********************************************************  
  SELECT "tBukv"  
  SCAN  
    lpValue=STRTRAN(lpValue, ALLTRIM(tBukv.cKir), ALLTRIM(tBukv.cLat))  
  endSCAN  
  USE IN "tBukv"  
  RETURN lpValue
  
*
* �������������� ������� ���� � ��������. ������� 2
*
FUNCTION translit
  LPARAMETERS _cString  
  _cString=UPPER(_cString)  
  LOCAL _cRetStr,_a,_b,a1,a2,a3,a4,b1,b2,b3,b4  
  a1="�"  
  a2="�"  
  a3="�"  
  a4="�"  
    
  b1="SH"  
  b2="CH"  
  b3="YA"  
  b4="CH"  
    
  _a="��������կԲ�������ƪ������޹"  
  _b="YCUKENGZHIFIVAPROLDJESMIT'BUN"  
    
  _cRetStr=CHRTRAN(_cString,_a,_b)  
  _cRetStr=STRTRAN(_cRetStr,a1,b1)  
  _cRetStr=STRTRAN(_cRetStr,a2,b2)  
  _cRetStr=STRTRAN(_cRetStr,a3,b3)  
  _cRetStr=STRTRAN(_cRetStr,a4,b4)  
  RETURN _cRetStr  

*
* Net Time ��� ������� ��� ��������� ������� � ������� � ���� �� ������ API NetRemoteTOD
* 08-10-2004
* http://www.foxclub.ru/sol/index.php?act=view&id=425
*
FUNCTION NetTime(m.cServerName)
*��� ������� ����������� � �������� \\, �������� \\server

PRIVATE m.lcServerName, m.lnBufferPointer, m.cSetDate, m.nHour, m.nMin, m.nSec, m.nmDiv, m.nDay, m.nMon, m.nYear, m.nRetValue
m.cSetDate=SET("DATE")
SET DATE GERMAN 
m.nRetValue={..::}
LOCAL ARRAY Arrdlls[1,3]
=ADLLS(Arrdlls)
IF ASCAN(Arrdlls,"NetRemoteTOD")=0
DECLARE INTEGER NetRemoteTOD IN NETAPI32 STRING ServerName,INTEGER @BufferPointer
ENDIF 
IF ASCAN(Arrdlls,"NetApiBufferFree")=0
DECLARE INTEGER NetApiBufferFree IN NETAPI32 INTEGER Pointer
ENDIF
IF ASCAN(Arrdlls,"RtlMoveMemory")=0
DECLARE RtlMoveMemory IN Kernel32 String@ dest, Integer src, Integer nsize
ENDIF 
IF ASCAN(Arrdlls,"lstrlenW")=0
DECLARE Integer lstrlenW IN Kernel32 Integer src
ENDIF 

m.cServerName=ALLTRIM(m.cServerName)
m.lcServerName = StrConv(StrConv(m.cServerName + Chr(0), 1), 5)
m.lnBufferPointer = 0
IF NetRemoteTOD(m.lcServerName,@m.lnBufferPointer)=0 && ��� ������!
m.nHour = GetMemoryDWORD(lnBufferPointer+8) &&����
m.nMin = GetMemoryDWORD(lnBufferPointer+12) &&������
m.nSec = GetMemoryDWORD(lnBufferPointer+16) &&�������
m.nmDiv = GetMemoryDWORD(lnBufferPointer+24) &&��������
m.nDay = GetMemoryDWORD(lnBufferPointer+32) &&����
m.nMon = GetMemoryDWORD(lnBufferPointer+36) &&�����
m.nYear = GetMemoryDWORD(lnBufferPointer+40) &&���
m.nRetValue=CTOT(PADL(ALLTRIM(STR(m.nDay)),2,"0")+"."+;
PADL(ALLTRIM(STR(m.nMon)),2,"0")+"."+;
PADL(ALLTRIM(STR(m.nYear)),4,"0")+" "+;
PADL(ALLTRIM(STR(m.nHour)),2,"0")+":"+;
PADL(ALLTRIM(STR(m.nMin)),2,"0")+":"+; 
PADL(ALLTRIM(STR(m.nsec)),2,"0"))-(60*m.nmDiv)
ENDIF
SET DATE &cSetDate
NetApiBufferFree(m.lnBufferPointer)
RETURN m.nRetValue


FUNCTION GetMemoryLPString
LParameter m.nPointer, m.nPlatform
Local m.cResult, m.nLPStr, m.nslen

m.nLPStr = GetMemoryDWORD(m.nPointer)

m.nslen = lstrlenW(m.nLPStr) * 2
m.cResult = Replicate(chr(0), m.nslen)
RtlMoveMemory(@m.cResult, m.nLPStr, m.nslen)
m.cResult = StrConv(StrConv(m.cResult, 6), 2)

RETURN m.cResult
ENDFUNC
*--
* Return Number from a pointer to DWORD
*--
FUNCTION GetMemoryDWORD
LParameter m.nPointer
Local m.cstrDWORD
m.cstrDWORD = Replicate(chr(0), 4)
RtlMoveMemory(@m.cstrDWORD, m.nPointer, 4)
RETURN str2long(m.cstrDWORD)
ENDFUNC

*-- 
* Convert a Number intto a binary Long Integer
*--
FUNCTION long2str
LParameter m.nLong

m.nLong = Int(m.nLong)
RETURN (chr(bitand(m.nLong,255)) + ;
chr(bitand(bitrshift(m.nLong, 8), 255)) + ;
chr(bitand(bitrshift(m.nLong, 16), 255)) + ;
chr(bitand(bitrshift(m.nLong, 24), 255)))
ENDFUNC
*-- 
* Convert a binary Long Integer into a Number
*--
FUNCTION str2long
LParameter m.cpLong
Return (Bitlshift(Asc(Substr(m.cpLong, 4, 1)), 24) + ;
Bitlshift(Asc(Substr(m.cpLong, 3, 1)), 16) + ;
Bitlshift(Asc(Substr(m.cpLong, 2, 1)), 8) + ;
Asc(Substr(m.cpLong, 1, 1)))
ENDFUNC
*
* End Function NetTime


FUNCTION ResizeGrid
LPARAMETERS oGrid as Grid
DECLARE INTEGER GetSystemMetrics IN user32.dll INTEGER nIndex  
oGrid.Width=MAX(GetSystemMetrics(2),GetSystemMetrics(3))+2  
oGrid.Height=MAX(GetSystemMetrics(2),GetSystemMetrics(3))+3
RETURN

FUNCTION UpdateReportCodePage
	MyFrx=GETFILE("FRX")
	IF EMPTY(MyFrx)
		RETURN .F.
	ENDIF
	Update &MyFrx set resoid=204
ENDFUNC


PROCEDURE GetGUID
DECLARE INTEGER CoCreateGuid ;  
    IN Ole32.dll ;  
    STRING @lcGUIDStruc  
  DECLARE INTEGER StringFromGUID2 ;  
    IN Ole32.dll ;  
    STRING cGUIDStruc, ;  
    STRING @cGUID, ;  
    LONG nSize
  
cStrucGUID=SPACE(16)  
cGUID=SPACE(80)  
nSize=40  
IF CoCreateGuid(@cStrucGUID) # 0  
	RETURN ""  
ENDIF  
IF StringFromGUID2(cStrucGUID,@cGuid,nSize) = 0  
	RETURN ""  
ENDIF  
RETURN STRCONV(LEFT(cGUID,76),6)  
ENDPROC

FUNCTION FlashWindow(nCount as Integer) as Integer
	#DEFINE FLASHW_STOP         0
	#DEFINE FLASHW_CAPTION      0x00000001
	#DEFINE FLASHW_TRAY         0x00000002
	#DEFINE FLASHW_ALL          BITOR(FLASHW_CAPTION, FLASHW_TRAY)
	#DEFINE FLASHW_TIMER        0x00000004
	#DEFINE FLASHW_TIMERNOFG    0x0000000C

	LOCAL lcFI as String
	DECLARE INTEGER FlashWindowEx IN WIN32API STRING pfwi
	lcFI = n2dword(20) + n2dword(_VFP.hWnd) + n2dword(FLASHW_ALL) + n2dword(nCount) + n2dword(0)
	
	#UNDEF FLASHW_STOP         
	#UNDEF FLASHW_CAPTION      
	#UNDEF FLASHW_TRAY         
	#UNDEF FLASHW_ALL          
	#UNDEF FLASHW_TIMER        
	#UNDEF FLASHW_TIMERNOFG    
	RETURN (FlashWindowEx(m.lcfi))
ENDFUNC

FUNCTION n2dword (tnNum)  
	lcRes = CHR(m.tnNum%256)  
	m.tnNum = INT(m.tnNum/256)  
	lcRes = m.lcRes + CHR(m.tnNum%256)  
	m.tnNum = INT(m.tnNum/256)  
	lcRes = m.lcRes + CHR(m.tnNum%256)  
	m.tnNum = INT(m.tnNum/256)  
	lcRes = m.lcRes + CHR(m.tnNum%256)  
RETURN m.lcRes

PROCEDURE LOC_CMonth  
  LPARAMETERS tuDateOrMonth, tuLCID, tcFormat  
 * �������������� ���� � ��������� ���� �� ��������� ����  
 * ���������:  
 * 1 - ������������� ����, ��������� ���������� ����-�����,  
 *  � ������ ����� ������.  
 *  �� ��������� ������������ ������� ��������� ����.  
 * 2 - LCID (1033 - ����������, 1049 - �������, 1059 - �����������),  
 *  ����� ��������� ��������� "S" � "D" ��� ��������������  
 *  ��������� � ���������������� ������.  
 *  �� ��������� ������������ ���������������� ������.  
 * 3 - ������, ��� ������� � MSDN. �������� ����� ������� ���:  
 *  MMMM - ������ �������� ������ (������������ �� ���������)  
 *  MMM - ����������� �������� ������  
 *  dddd - ������ �������� ��� ������  
 *  ddd - ����������� �������� ��� ������  
 *  
 * ������������ �:  
 *  
 * ��������: SWord()  
 *  
    
  #DEFINE LOCALE_USER_DEFAULT   0x00000400  
  #DEFINE LOCALE_SYSTEM_DEFAULT 0x00000800  
  LOCAL ldDate, lnLCID, lcFormat, lcStruct, lcRetVal, lnLen  
  DO CASE  
  CASE PCOUNT() = 0 OR ;  
    EMPTY(m.tuDateOrMonth)  
   ldDate = DATE()  
  CASE VARTYPE(m.tuDateOrMonth) = "D"  
   ldDate = m.tuDateOrMonth  
  CASE VARTYPE(m.tuDateOrMonth) = "T"  
   ldDate = TTOD(m.tuDateOrMonth)  
  CASE VARTYPE(m.tuDateOrMonth) = "N"  
   ldDate = DATE(2000, m.tuDateOrMonth, 1)  
  OTHERWISE  
   ERROR 11  
   RETURN ""  
  ENDCASE  
  DO CASE  
  CASE PCOUNT() < 2 OR ;  
    (VARTYPE(m.tuLCID) = "C" AND ;  
    UPPER(LEFT(m.tuLCID, 1)) = "U") OR ;  
    EMPTY(m.tuLCID)  
   lnLCID = LOCALE_USER_DEFAULT  
  CASE VARTYPE(m.tuLCID) = "N"  
   lnLCID = m.tuLCID  
  CASE VARTYPE(m.tuLCID) = "C" AND ;  
    UPPER(LEFT(m.tuLCID, 1)) = "S"  
   lnLCID = LOCALE_SYSTEM_DEFAULT  
  OTHERWISE  
   ERROR 11  
   RETURN ""  
  ENDCASE  
  DO CASE  
  CASE PCOUNT() < 3 OR ;  
    EMPTY(m.tcFormat)  
   lcFormat = "MMMM"  
  CASE VARTYPE(m.tcFormat) = "C"  
   lcFormat = m.tcFormat  
  OTHERWISE  
   ERROR 11  
   RETURN ""  
  ENDCASE  
  lcStruct = ;  
   (SWord(YEAR(m.ldDate)) + ;  
   SWord(MONTH(m.ldDate)) + ;  
   SWord(1) + ;  
   SWord(DAY(m.ldDate)) + ;  
   SWord(0) + ;  
   Sword(0) + ;  
   SWord(0) + ;  
   SWord(0) + ;  
   SWord(0))  
  DECLARE INTEGER GetDateFormat IN WIN32API ;  
   INTEGER LCID, ;  
   INTEGER dwFlags, ;  
   STRING SystemDate, ;  
   STRING lpFormat, ;  
   STRING@ lpDateStr, ;  
   INTEGER lenDate  
  lcRetVal = REPLICATE(CHR(0), 128)  
  lnLen = GetDateFormat(m.lnLCID, 0, m.lcStruct, m.lcFormat, @m.lcRetVal, LEN(m.lcRetVal))  
  RETURN LEFT(m.lcRetVal, MAX(m.lnLen - 1, 0))  
  ENDPROC  
    
  PROCEDURE SWord  
  LPARAMETERS tnInt  
  RETURN CHR(m.tnInt % 256) + CHR(INT(m.tnInt / 256))  
  ENDPROC
  
  FUNCTION Unicode2KOI8R
  Lparameters nCurrentCodePage, nNewCodePage, cString   
    
  Declare Integer IsValidCodePage in WIN32API ;  
  	integer nCodePage  
    
  Declare Integer MultiByteToWideChar in WIN32API ;  
  	integer CodePage,;  
  	integer Flags,;  
  	string MultyByteStr,;  
  	integer MultiByteStrLen,;  
  	string @ WideCharStr,;  
  	integer WideCharStrLen  
    
  Declare Integer WideCharToMultiByte in WIN32API ;  
  	integer CodePage,;  
  	integer Flags,;  
  	string MultyByteStr,;  
  	integer MultiByteStrLen,;  
  	string @ WideCharStr,;  
  	integer WideCharStrLen,;  
  	integer ,;  
  	integer   
  	  
  If IsValidCodePage(nCurrentCodePage) = 0  
  	Error 1914  
  	Return ""  
  EndIf  
  	  
  If IsValidCodePage(nNewCodePage) = 0  
  	Error 1914  
  	Return ""  
  EndIf  
  	  
  Local WideCharBuf, MultiByteBuf  
  WideCharBuf=Replicate(Chr(0),Len(cString)*2)  
  MultiByteBuf=Replicate(Chr(0),Len(cString))  
    
  MultiByteToWideChar;  
  	(nCurrentCodePage;  
  	,0;  
  	,cString;  
  	,Len(cString);  
  	,@WideCharBuf;  
  	,Len(WideCharBuf))  
    
  WideCharToMultiByte;  
  	(nNewCodePage;  
  	,0;  
  	,WideCharBuf;  
  	,Len(WideCharBuf);  
  	,@MultiByteBuf;  
  	,Len(MultiByteBuf);  
  	,0,0)  
  	  
  Return MultiByteBuf
  

function HRUN
  LPARAM lcFileName,lcWorkDir,lcOperation,cParameters,nWinHandle,lnShow  
    
  lcFileName=ALLTRIM(lcFileName)  
  lcWorkDir=IIF(TYPE("lcWorkDir")="C",ALLTRIM(lcWorkDir),"")+chr(0)  
  lcOperation=IIF(TYPE("lcOperation")="C" AND NOT EMPTY(lcOperation),ALLTRIM(lcOperation),"Open")  
  lnShow=IIF(TYPE("lnShow")="N",lnShow,0) && 0 - ��������� ����  
  cParameters=IIF(TYPE("cParameters")="C",alltrim(cParameters),"")+chr(0) && ��������� (�����)  
  nWinHandle=IIF(TYPE("nWinHandle")="N",nWinHandle,0) && ��. api FindWindow  
    
  DECLARE INTEGER ShellExecute ;  
  IN SHELL32.DLL ;  
  INTEGER nWinHandle,;  
  STRING cOperation,;  
  STRING cFileName,;  
  STRING cParameters,;  
  STRING cDirectory,;  
  INTEGER nShowWindow  
  RETURN ShellExecute(nWinHandle,lcOperation,lcFilename,@cParameters,@lcWorkDir,lnShow)
  
*  lpr='a -sfx -s -ag -hpalexg  -m5 -md1024  -inul -c- -idp '+g_patarch+' "'+g_pathist+'"'  
*  =hrun('rar.exe',g_pathrar,,lpr)  
*        =MESSAGEBOX ("�������� ������ ���������")

*
* ������� ����� ������ ������
*
FUNCTION MakeNewDataSession
	oSes=CREATEOBJECT("session")  
	SET DATASESSION TO oSes.DataSession  
	RETURN (oSes)
ENDFUNC

* ������� ��������� �������
FUNCTION ProcessEvents
	LOCAL cWM  
	ndLlscnt = ADLLS(ad)
	IF ndLlscnt>0
		IF ASCAN(ad, "PeekMessage")=0
			DECLARE integer PeekMessage IN WIN32API AS WinAPI_PeekMessage string@, integer, integer, integer, integer  
		ENDIF
		IF ASCAN(ad, "TranslateMessage")=0
			DECLARE integer TranslateMessage IN WIN32API  AS WinAPI_TranslateMessage string@
		ENDIF
		IF ASCAN(ad, "DispatchMessage")=0
			DECLARE integer DispatchMessage IN WIN32API  AS WinAPI_DispatchMessage string@
		ENDIF
	ELSE
	* �������
	DECLARE integer PeekMessage IN WIN32API AS WinAPI_PeekMessage string@, integer, integer, integer, integer  
	DECLARE integer TranslateMessage IN WIN32API  AS WinAPI_TranslateMessage string@  
	DECLARE integer DispatchMessage IN WIN32API  AS WinAPI_DispatchMessage string@
ENDIF

	IF _vfp.AutoYield   
		* � ���� ������ ����� ���������� �� VFP  
		DOEVENTS  
	ELSE  
		* ������� �� VFP ��� ������ - ���������� ��������� ������� ���������	  
		m.cWM = Space(28)  
		DO WHILE ( WinAPI_PeekMessage( @cWM, 0, 0, 0, 1 ) <> 0 )  
			WinAPI_TranslateMessage( @cWM )  
			WinAPI_DispatchMessage( @cWM )  
		ENDDO  
		* � ������ �� ���������� ������� VFP ���� ���������  
		DOEVENTS  
	ENDIF  
ENDFUNC


* ������� ������ SQL �������� � �������������� SQLDMO
FUNCTION EnumerateSQLSERVER
	CLEAR
	oSQLApp = CREATEOBJECT('SQLDMO.Application')
	oSqlServerList = oSQLApp.ListAvailableSQLServers()
	FOR i = 1 TO oSqlServerList.COUNT
		=oSqlServerList.ITEM(i)
	ENDFOR
ENDFUNC

*
* ����������� ������ EXE-�����
*
FUNCTION VersIoninfo
LPARAMETERS lcExeName as Character
IF !FILE(lcExeName)
	RETURN ("")
ENDIF
LOCAL aFiles as Variant
DIMENSION aFiles[1]
AGETFILEVERSION(aFiles, lcExeName)
IF ALEN(aFiles)=15
	RETURN (aFiles[11])
ENDIF
RETURN ("")
ENDFUNC


*
* �������� ��� �������
*  tnHWND - ���������� ����
*  tcFrom - ������������ �������� ������ (����� ��������� �������� � "?")  
*  tcTo   - ������������ �������������� ������ (�����)  
*  tnOper - ���������� ��� ��������.  
*	     tnOper=1 - ����������� �����  
*	     tnOper=2 - ���������� �����  
*	     tnOper=3 - ������� �����  
*	     tnOper=4 - ������������� ����  
*  
FUNCTION FileOperation  
  LPARAMETER tnHWND, tcFrom, tcTo, tnOper  

  LOCAL lcSHFO, lcFrom, lnLenFrom, lcTo, lnLenTo, hGlobalFrom, hGlobalTo  
  LOCAL lnFlag, lnReturn  
  DECLARE Long SHFileOperation IN Shell32.dll String @
  DECLARE Long DeleteFile IN kernel32.dll String @
  
  DECLARE Long GlobalAlloc IN WIN32API Long, Long  
  DECLARE Long GlobalFree IN WIN32API Long
 *  
 * �������� ����������� ��������� � ���������� lcSHFO  
 * 
 	*_screen.HWnd
  * lcSHFO = BINTOC(thisform.HWnd, '4RS')      && ���� hwnd ���������  
  lcSHFO = BINTOC(tnHWND, '4RS')      && ���� hwnd ���������
  lcSHFO = lcSHFO + BINTOC(tnOper, '4RS')    && ���� wFunc - ��� ��������   
 *  
 * ��������� ������������ �������� ������  
 *  
  tcFrom = tcFrom + CHR(0) + CHR(0)             && ���������� ����  
  lnLenFrom = LEN(tcFrom)                       && ����� �������� ������  
  hGlobalFrom = GlobalAlloc(0x0040, lnLenFrom)  && �������� ��� ��  
                                                && ���� ������  
  SYS(2600, hGlobalFrom, lnLenFrom, tcFrom)     && � �������� ���� ������  
  lcSHFO = lcSHFO + BINTOC(hGlobalFrom, '4RS')  && ���� pFrom  
 *  
 * ��������� ������������ �������������� ������  
 *  
  IF tnOper = 3  
      lcSHFO = lcSHFO + BINTOC(0, '4RS') && ���� �������� ��������  
  ELSE  
      tcTo = tcTo + CHR(0) + CHR(0)      && ���������� ����  
      lnLenTo = LEN(tcTo)                && ����� �������������� ������  
      hGlobalTo = GlobalAlloc(0x0040, lnLenTo)   && �������� ��� ��  
                                                 && ���� ������  
      SYS(2600, hGlobalTo, lnLenTo, tcTo)        && � �������� ���� ������  
      lcSHFO = lcSHFO + BINTOC(hGlobalTo, '4RS') && ���� pTo ���������  
  ENDIF  
  lnFlag = 8 + 512  
  lcSHFO = lcSHFO + BINTOC(lnFlag, '2RS') && ���� fFlags ���������  
  lcSHFO = lcSHFO + REPLICATE(CHR(0), 12) && ��������� 3 ���� ���������  
 *  
 * ��������� ��������  
 *
 IF tnOper = 3
 	lnReturn = DeleteFile(@tcFrom)
 	IF lnReturn > 0
 		lnReturn = 0
 	ELSE
 		lnReturn = -1
 	ENDIF
 ELSE
 	lnReturn = SHFileOperation(@lcSHFO)  
 ENDIF
 *  
 * ���������� ������ Windows  
 *  
  GlobalFree(hGlobalFrom)  
  IF tnOper != 3  
      GlobalFree(hGlobalTo)  
  ENDIF  
 *  
 * ���� lnReturn = 0, �� �������� ��������� �������  
 *  
  IF lnReturn != 0  
      RETURN .f.  
  ENDIF  
  RETURN .t.
  
  *****************************************************************************************
* FUNCTION : RECURSE
* AUTHOR : MICHAEL REYNOLDS
* DESCRIPTION : Good for performing file processing throughout an entire directory tree.
* The function, RECURSE(), is called with the full path of a directory.
* RECURSE() will then read all files and directories in the path.
* A function can be called to process files that it finds.
* Plus, the function calls itself to process any sub-directories.
* http://fox.wikis.com/wc.dll?Wiki~RecursiveDirectoryProcessing~VFP
*****************************************************************************************
FUNCTION Recurse(pcDir)
LOCAL lnPtr, lnFileCount, laFileList, lcDir, lcFile
CHDIR (pcDir)
DIMENSION laFileList[1]
*--- Read the chosen directory.
lnFileCount = ADIR(laFileList, '*.*', 'D')
FOR lnPtr = 1 TO lnFileCount
   IF 'D' $ laFileList[lnPtr, 5]
   *--- Get directory name.
   lcDir = laFileList[lnPtr, 1]
   *--- Ignore current and parent directory pointers.
      IF lcDir != '.' AND lcDir != '..'
      *--- Call this routine again.
         Recurse(lcDir)
      ENDIF
   ELSE
      *--- Get the file name.
      lcFile = LOWER(FULLPATH(laFileList[lnPtr, 1]))
      *--- Insert into cursor if .EXE, .ICO or .DLL
      IF INLIST(LOWER(JUSTEXT(lcFile)),"dll","exe","ico")
         INSERT INTO Recursive (cFile) VALUES (lcFile)
      ENDIF
   ENDIF
ENDFOR
*--- Move back to parent directory.
CHDIR ..
ENDFUNC

*
* ���� ������������ ���
*
DEFINE CLASS FIO as Custom
	HIDDEN lcFullName
	HIDDEN lcLastName
	HIDDEN lcFirstName
	HIDDEN lcSubName
	
	lcFullName = ""
	lcLastName = ""
	lcFirstName = ""
	lcSubName = ""
	
	PROCEDURE Init
		LPARAMETERS lcFio as Character
		This.lcFullName = lcFio
		This.lcLastName = DIV_FIO(lcFio,1)
		This.lcFirstName = DIV_FIO(lcFio,2)
		This.lcSubName = DIV_FIO(lcFio,3)
	ENDPROC
	
	PROCEDURE LastName
		RETURN (This.lcLastName)
	ENDPROC
	PROCEDURE FirstName
		RETURN (This.lcFirstName)
	ENDPROC
	PROCEDURE SubName
		RETURN (This.lcSubName)
	ENDPROC
ENDDEFINE

*
* ������� ������� �� ��������� �����
*
* lcFio - ������� ��� ��������
* liMode - 1-��� 2-��� 3-��������
*
FUNCTION DIV_FIO(lcFio as Character, liMode as Integer)
IF PCOUNT()=0
	RETURN []
ENDIF

a1 = SUBSTR(lcFio, 1, AT(" ",lcFio))
lcFio = SUBSTR(lcFio, AT(" ",lcFio)+1)

a2 = SUBSTR(lcFio,1, AT(" ",lcFio))

lcFio = SUBSTR(lcFio, AT(" ",lcFio)+1)

a3 = SUBSTR(lcFio,1)

DO CASE
CASE liMode=1
	RETURN (a1)
CASE liMode=2
	RETURN (a2)
CASE liMode=3
	RETURN (a3)

OTHERWISE
	return(a1+" "+a2+" "+a3)
ENDCASE
RETURN []
ENDFUNC

********************************************************************************************
* ����� �������� Visual FoxPro
*
* ��������� ����� �� 0 �� 999'999'999'999,99
* ���������� ��������� ������ � ������ � ������
*
* ���� ����� �������������, �� ����� ������ �����.
* �� ������� �� ������ ����������, ����� ������ �������� �������� ��� ������,
* � ������ "������" � "������" ����������� ��������������� �����.
*
* ����������� �� �������� ��� ������������� ������ � ����� �� �����������.
********************************************************************************************
* nValuta = 0 - ������ �� ��������
*         = 1 - �����
*         = 2 - �������
********************************************************************************************
FUNCTION Sum_Str
PARAMETERS nSum, nValuta
LOCAL cRet,AswCounter,AswS,Asw,Asw1

cRet=""
nSum=iif(empty(nSum),0,abs(nSum))

AswS=str(nSum,15,2)
FOR AswCounter=1 to 10 STEP 3
Asw=substr(AswS,AswCounter,3)
if Asw#space(3)
Asw1=""
do case
case substr(Asw,1,1)="1"
Asw1="��� "
case substr(Asw,1,1)="2"
Asw1="������ "
case substr(Asw,1,1)="3"
Asw1="������ "
case substr(Asw,1,1)="4"
Asw1="��������� "
case substr(Asw,1,1)="5"
Asw1="������� "
case substr(Asw,1,1)="6"
Asw1="�������� "
case substr(Asw,1,1)="7"
Asw1="������� "
case substr(Asw,1,1)="8"
Asw1="��������� "
case substr(Asw,1,1)="9"
Asw1="��������� "
endcase
cRet=cRet+Asw1
Asw1=""
if substr(Asw,2,1)="1"
do case
case substr(Asw,3,1)="0"
Asw1="������ "
case substr(Asw,3,1)="1"
Asw1="����������� "
case substr(Asw,3,1)="2"
Asw1="���������� "
case substr(Asw,3,1)="3"
Asw1="���������� "
case substr(Asw,3,1)="4"
Asw1="������������ "
case substr(Asw,3,1)="5"
Asw1="���������� "
case substr(Asw,3,1)="6"
Asw1="����������� "
case substr(Asw,3,1)="7"
Asw1="���������� "
case substr(Asw,3,1)="8"
Asw1="������������ "
case substr(Asw,3,1)="9"
Asw1="������������ "
endcase
cRet=cRet+Asw1
do case
case AswCounter=1
cRet=cRet+"���������� "
case AswCounter=4
cRet=cRet+"��������� "
case AswCounter=7
cRet=cRet+"����� "
case AswCounter=10
	DO CASE
	CASE nValuta = 0
		cRet=cRet+" "
	CASE nValuta = 1
		cRet=cRet+"������ "
	CASE nValuta = 2
		cRet=cRet+"USD "
	ENDCASE
endcase
else
do case
case substr(Asw,2,1)="2"
Asw1="�������� "
case substr(Asw,2,1)="3"
Asw1="�������� "
case substr(Asw,2,1)="4"
Asw1="����� "
case substr(Asw,2,1)="5"
Asw1="��������� "
case substr(Asw,2,1)="6"
Asw1="���������� "
case substr(Asw,2,1)="7"
Asw1="��������� "
case substr(Asw,2,1)="8"
Asw1="����������� "
case substr(Asw,2,1)="9"
Asw1="��������� "
endcase
cRet=cRet+Asw1
Asw1=""
do case
case substr(Asw,1,3)="000"
Asw1=iif(AswCounter=10,"������ ","")
case substr(Asw,3,1)="0" .and. substr(Asw,1,3)#"000"
Asw1=iif(AswCounter=7,"����� ", ;
iif(AswCounter=1,"���������� ",iif(AswCounter=4,"��������� ","������ ")))
case substr(Asw,3,1)="1"
Asw1=iif(AswCounter=7,"���� ������ ", ;
"���� "+iif(AswCounter=1,"�������� ",iif(AswCounter=4,"������� ","����� ")))
case substr(Asw,3,1)="2"
Asw1=iif(AswCounter=7,"��� ������ ", ;
"��� "+iif(AswCounter=1,"��������� ",iif(AswCounter=4,"�������� ","����� ")))
case substr(Asw,3,1)="3"
Asw1=iif(AswCounter=7,"��� ������ ", ;
"��� "+iif(AswCounter=1,"��������� ",iif(AswCounter=4,"�������� ","����� ")))
case substr(Asw,3,1)="4"
Asw1=iif(AswCounter=7,"������ ������ ", ;
"������ "+iif(AswCounter=1,"��������� ",iif(AswCounter=4,"�������� ","����� ")))
case substr(Asw,3,1)="5"
Asw1=iif(AswCounter=7,"���� ����� ", ;
"���� "+iif(AswCounter=1,"���������� ",iif(AswCounter=4,"��������� ","������ ")))
case substr(Asw,3,1)="6"
Asw1=iif(AswCounter=7,"����� ����� ", ;
"����� "+iif(AswCounter=1,"���������� ",iif(AswCounter=4,"��������� ","������ ")))
case substr(Asw,3,1)="7"
Asw1=iif(AswCounter=7,"���� ����� ", ;
"���� "+iif(AswCounter=1,"���������� ",iif(AswCounter=4,"��������� ","������ ")))
case substr(Asw,3,1)="8"
Asw1=iif(AswCounter=7,"������ ����� ", ;
"������ "+iif(AswCounter=1,"���������� ",iif(AswCounter=4,"��������� ","������ ")))
case substr(Asw,3,1)="9"
Asw1=iif(AswCounter=7,"������ ����� ", ;
"������ "+iif(AswCounter=1,"���������� ",iif(AswCounter=4,"��������� ","������ ")))
endcase
cRet=cRet+Asw1
endif
endif
ENDFOR
AswS=substr(AswS,14,2)

if substr(AswS,1,1)="1" .or. substr(AswS,2,1)="0"
cRet=cRet+AswS+" ������"
else
do case
case substr(AswS,2,1)>"4"
cRet=cRet+AswS+" ������"
case substr(AswS,2,1)="1"
cRet=cRet+AswS+" �������"
case substr(AswS,2,1)>"1" .and. substr(AswS,2,1)<"5"
cRet=cRet+AswS+" �������"
endcase
endif
cRet=upper(substr(cRet,1,1))+substr(cRet,2)
if nSum<1
cRet="���� "+cRet
endif
return cRet


* Function GetProxySettings  
* ��������: forum.foxclub.ru  
* �� ������ ����: www.tech-archive.net  
*
FUNCTION GetProxySettings
  #define INTERNET_OPTION_PROXY 38  
    
  LOCAL lProxyAndPort, buffer, cbBuffer  
    
  lProxyAndPort = Space(0)  
    
  DECLARE INTEGER InternetQueryOption IN wininet;  
      INTEGER   hInternet,;  
      INTEGER   lOption,;  
      STRING  @ sBuffer,;  
      INTEGER @ lBufferLength  
    
  m.buffer = SPACE(4096)  
  m.cbBuffer = LEN(m.buffer)  
    
  ret = InternetQueryOption(0, INTERNET_OPTION_PROXY, @buffer, cbBuffer)  
  If ret = 1  
  	buffer = SubStr(buffer, 13)  
  	lProxyAndPort = Left(buffer, AT(Chr(0),buffer)-1)  
  EndIf  
    
  CLEAR  
  ? "Proxy � ����: "+lProxyAndPort  
    
  RETURN
ENDFUNC

*
* Function: INPUT_MSG(m.cInputPrompt, m.cDialogCaption, m.cDefaultValue, m.cInputFormat)
*
* m.cInputPrompt - ����� �������������� ��� ����� ��� �����
* m.cDialogCaption - ��������� ����
* m.cDefaultValue - �������� �� ���������
* m.cInputFormat - ������, "K","D","T","999,99" ��� ����� ������
*
FUNCTION INPUT_MSG
LPARAMETERS m.cInputPrompt, m.cDialogCaption, m.cDefaultValue, m.cInputFormat
LOCAL oForm as input_msg OF D:\fox\vcl\__form, RetValue
oForm = NEWOBJECT("input_msg","D:\fox\vcl\__form","",m.cInputPrompt,m.cDialogCaption,m.cDefaultValue, m.cInputFormat)
oForm.Show(1)
IF ISNULL(oForm)
	RetValue = NULL
ELSE
	RetValue = oForm.returnvalue
ENDIF
RETURN (RetValue)
ENDFUNC

FUNCTION Speeling  
  PARAMETER nSumma  
  PRIVATE cSumma  
   * k - �������  
    cSumma = TRANSFORM(M.nSumma,'9,9,,9,,,,,,9,9,,9,,,,,9,9,,9,,,,9,9,,9,,,.99')+'k'  

   * t - ������; m - ��������; M - ���������  
    cSumma = STRTRAN(M.cSumma, ',,,,,,', 'eM')  
    cSumma = STRTRAN(M.cSumma, ',,,,,',  'em')  
    cSumma = STRTRAN(M.cSumma, ',,,,',   'et')  
    
   * e - �������; d - �������; c - �����  
    cSumma = STRTRAN(M.cSumma, ',,,', 'e')  
    cSumma = STRTRAN(M.cSumma, ',,',  'd')  
    cSumma = STRTRAN(M.cSumma, ',',   'c')  
    
    cSumma = STRTRAN(M.cSumma, '0c0d0et', '')  
    cSumma = STRTRAN(M.cSumma, '0c0d0em', '')  
    cSumma = STRTRAN(M.cSumma, '0c0d0eM', '')  
    
    cSumma = STRTRAN(M.cSumma, '0c', '')  
    cSumma = STRTRAN(M.cSumma, '1c', '��� ')  
    cSumma = STRTRAN(M.cSumma, '2c', '������ ')  
    cSumma = STRTRAN(M.cSumma, '3c', '������ ')  
    cSumma = STRTRAN(M.cSumma, '4c', '��������� ')  
    cSumma = STRTRAN(M.cSumma, '5c', '������� ')  
    cSumma = STRTRAN(M.cSumma, '6c', '�������� ')  
    cSumma = STRTRAN(M.cSumma, '7c', '������� ')  
    cSumma = STRTRAN(M.cSumma, '8c', '��������� ')  
    cSumma = STRTRAN(M.cSumma, '9c', '��������� ')  
    
    cSumma = STRTRAN(M.cSumma, '1d0e', '������ ')  
    cSumma = STRTRAN(M.cSumma, '1d1e', '����������� ')  
    cSumma = STRTRAN(M.cSumma, '1d2e', '���������� ')  
    cSumma = STRTRAN(M.cSumma, '1d3e', '���������� ')  
    cSumma = STRTRAN(M.cSumma, '1d4e', '������������ ')  
    cSumma = STRTRAN(M.cSumma, '1d5e', '���������� ')  
    cSumma = STRTRAN(M.cSumma, '1d6e', '����������� ')  
    cSumma = STRTRAN(M.cSumma, '1d7e', '����������� ')  
    cSumma = STRTRAN(M.cSumma, '1d8e', '������������ ')  
    cSumma = STRTRAN(M.cSumma, '1d9e', '������������ ')  
 
    cSumma = STRTRAN(M.cSumma, '0d', '')  
    cSumma = STRTRAN(M.cSumma, '2d', '�������� ')  
    cSumma = STRTRAN(M.cSumma, '3d', '�������� ')  
    cSumma = STRTRAN(M.cSumma, '4d', '����� ')  
    cSumma = STRTRAN(M.cSumma, '5d', '��������� ')  
    cSumma = STRTRAN(M.cSumma, '6d', '���������� ')  
    cSumma = STRTRAN(M.cSumma, '7d', '��������� ')  
    cSumma = STRTRAN(M.cSumma, '8d', '����������� ')  
    cSumma = STRTRAN(M.cSumma, '9d', '��������� ')  
    
    cSumma = STRTRAN(M.cSumma, '0e', '')  
    cSumma = STRTRAN(M.cSumma, '5e', '���� ')  
    cSumma = STRTRAN(M.cSumma, '6e', '����� ')  
    cSumma = STRTRAN(M.cSumma, '7e', '���� ')  
    cSumma = STRTRAN(M.cSumma, '8e', '������ ')  
    cSumma = STRTRAN(M.cSumma, '9e', '������ ')  

***************************************************************
    cSumma = STRTRAN(M.cSumma, '1e.', '���� ')  
    cSumma = STRTRAN(M.cSumma, '2e.', '��� ')  
    cSumma = STRTRAN(M.cSumma, '3e.', '��� ')  
    cSumma = STRTRAN(M.cSumma, '4e.', '������ ')  
    cSumma = STRTRAN(M.cSumma, '1et', '���� ������ ')  
    cSumma = STRTRAN(M.cSumma, '2et', '��� ������ ')  
    cSumma = STRTRAN(M.cSumma, '3et', '��� ������ ')  
    cSumma = STRTRAN(M.cSumma, '4et', '������ ������ ')  
    cSumma = STRTRAN(M.cSumma, '1em', '���� ������� ')  
    cSumma = STRTRAN(M.cSumma, '2em', '��� �������� ')  
    cSumma = STRTRAN(M.cSumma, '3em', '��� �������� ')  
    cSumma = STRTRAN(M.cSumma, '4em', '������ �������� ')  
    cSumma = STRTRAN(M.cSumma, '1eM', '���� �������� ')  
    cSumma = STRTRAN(M.cSumma, '2eM', '��� ��������� ')  
    cSumma = STRTRAN(M.cSumma, '3eM', '��� ��������� ')  
    cSumma = STRTRAN(M.cSumma, '4eM', '������ ��������� ')  
    
    cSumma = STRTRAN(M.cSumma, '11k', '����������� ')  
    cSumma = STRTRAN(M.cSumma, '12k', '���������� ')  
    cSumma = STRTRAN(M.cSumma, '13k', '���������� ')  
    cSumma = STRTRAN(M.cSumma, '14k', '������������ ')  
    cSumma = STRTRAN(M.cSumma, '1k', '���� ')  
    cSumma = STRTRAN(M.cSumma, '2k', '��� ')  
    cSumma = STRTRAN(M.cSumma, '3k', '��� ')  
    cSumma = STRTRAN(M.cSumma, '4k', '������ ')  
    
    cSumma = STRTRAN(M.cSumma, '.', '')  
    cSumma = STRTRAN(M.cSumma, 't', '����� ')  
    cSumma = STRTRAN(M.cSumma, 'm', '��������� ')  
    cSumma = STRTRAN(M.cSumma, 'M', '���������� ')  
    cSumma = STRTRAN(M.cSumma, 'k', '')  
    cSumma = STRTRAN(M.cSumma, ' 00', '')  
***************************************************************
    
*!*	    cSumma = STRTRAN(M.cSumma, '1e.', '���� ����� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2e.', '��� ����� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3e.', '��� ����� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4e.', '������ ����� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '1et', '���� ������ ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2et', '��� ������ ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3et', '��� ������ ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4et', '������ ������ ')  
*!*	    cSumma = STRTRAN(M.cSumma, '1em', '���� ������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2em', '��� �������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3em', '��� �������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4em', '������ �������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '1eM', '���� �������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2eM', '��� ��������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3eM', '��� ��������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4eM', '������ ��������� ')  
*!*	    
*!*	    cSumma = STRTRAN(M.cSumma, '11k', '11 ������')  
*!*	    cSumma = STRTRAN(M.cSumma, '12k', '12 ������')  
*!*	    cSumma = STRTRAN(M.cSumma, '13k', '13 ������')  
*!*	    cSumma = STRTRAN(M.cSumma, '14k', '14 ������')  
*!*	    cSumma = STRTRAN(M.cSumma, '1k', '1 �������')  
*!*	    cSumma = STRTRAN(M.cSumma, '2k', '2 �������')  
*!*	    cSumma = STRTRAN(M.cSumma, '3k', '3 �������')  
*!*	    cSumma = STRTRAN(M.cSumma, '4k', '4 �������')  
*!*	    
*!*	    cSumma = STRTRAN(M.cSumma, '.', '������ ')  
*!*	    cSumma = STRTRAN(M.cSumma, 't', '����� ')  
*!*	    cSumma = STRTRAN(M.cSumma, 'm', '��������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, 'M', '���������� ')  
*!*	    cSumma = STRTRAN(M.cSumma, 'k', ' ������')  
    m.cSumma=allt(IIF(M.nSumma < 10**12, M.cSumma, ALLTRIM(STR(M.nSumma,20,2))))  
    m.cSumma=upper(left(m.cSumma,1))+lower(substr(m.cSumma,2))  
    
  RETURN m.cSumma 

FUNCTION SummaPropis
LPARAMETERS eSumma  
  LOCAL v,v123,n1,n2,n12,n3,e10m,e10z,e19,e90,e900,m,c123,r3,cSumma,nPoint,cPoint,lni,cBig,cSmall,cDrob  
  m.cPoint=SET("Point")  
  m.nPoint=IIF(VARTYPE(m.eSumma)='N',0,AT(m.cPoint,m.eSumma))  
  m.r3=CEILING(LEN(LTRIM(ICASE(VARTYPE(m.eSumma)='N',STR(INT(m.eSumma),24),VARTYPE(m.eSumma)='C',IIF(m.nPoint=0,m.eSumma,LEFT(m.eSumma,m.nPoint-1)),'')))/3)  
  m.cSumma=IIF(VARTYPE(m.eSumma)='C',PADL(IIF(m.nPoint=0,m.eSumma+'.00',PADR(m.eSumma,m.r3*3+2,'0')),m.r3*3+3),TRANSFORM(m.eSumma,'@l '+REPLICATE('9',m.r3*3)+'.99'))  
  m.nPoint=AT(m.cPoint,m.cSumma)    
  m.e900='���, ������, ������, ���������, �������, ��������, �������, ���������, ���������,'  
  m.e90='������, ��������, ��������, �����, ���������, ����������, ���������, �����������, ���������,'  
  m.e19='�����������, ����������, ����������, ������������, ����������, �����������, ����������, ������������, ������������,'  
  m.e10m='����, ���, ���, ������, ����, �����, ����, ������, ������,'  
  m.e10z='����, ���, ���, ������, ����, �����, ����, ������, ������,'  
  m.m='����� ������� �������� �������� ����������� ����������� ����������� ���������� ��������� ��������� ���������'  
  m.v=''  
  FOR lni=m.r3 TO 1 STEP -1  
  	m.c123=SUBSTR(m.cSumma,m.nPoint-3*lni,3)  
  	m.n12=IIF(BETWEEN(RIGHT(m.c123,2),'11','19'),VAL(RIGHT(m.c123,1)),0)  
  	m.n2=IIF(m.n12=0,VAL(SUBSTR(m.c123,2,1)),0)  
  	m.n1=IIF(m.n12=0 and m.n2#1,VAL(RIGHT(m.c123,1)),0)  
  	m.n3=VAL(LEFT(m.c123,1))  
  	m.v123=GETWORDNUM(m.e900,n3)+GETWORDNUM(m.e19,m.n12)+GETWORDNUM(m.e90,m.n2)+GETWORDNUM(IIF(lni=2,m.e10z,m.e10m),m.n1)  
  	m.v=m.v+m.v123+IIF(lni=1,'',IIF(VAL(m.c123)>0,GETWORDNUM(m.m,lni-1)+IIF(m.n1=1,IIF(lni=2,'a',''),IIF(BETWEEN(m.n1,2,4),IIF(lni=2,'�','�'),IIF(lni=2,'','��')))+',',''))  
  ENDFOR  
  m.v=IIF(EMPTY(m.v),'���� ',m.v)  
 *RETURN CHRTRAN(PROPER(m.v),',',' ')+SUBSTR(m.cSumma,m.nPoint+1,2)  
  m.cDrob=SUBSTR(m.cSumma,m.nPoint+1,2)  
  m.cBig='����'+ICASE(m.n12>0 OR m.n1=0 OR BETWEEN(m.n1,5,9),'��',m.n12=0 AND BETWEEN(m.n1,2,4),'�','�')+' '  
  m.cSmall=' ����'+ICASE(LEFT(m.cDrob,1)='1' OR RIGHT(m.cDrob,1)='0' OR BETWEEN(RIGHT(m.cDrob,1),'5','9'),'��',BETWEEN(RIGHT(m.cDrob,1),'2','4'),'���','���')  
  RETURN CHRTRAN(PROPER(m.v),',',' ') + m.cBig + m.cDrob + m.cSmall
ENDFUNC

FUNCTION ErrorHandler(nError,cMethod,nLine)
	DO WHILE TXNLEVEL( )>0
		ROLLBACK
	ENDDO

	DO CASE
	CASE nError = 1429
		MESSAGEBOX("#1429.SQL Server �� ����������, ��� ������ ��������..."+CHR(13)+"���������� � ���������� �������������� ��� ��������� �������"+CHR(13)+MESSAGE(), 48, "������ ����������� � SQL-�������")
		ON ERROR
		QUIT

	CASE nError = 1435
		MESSAGEBOX("#1435.������ � ��������� ������"+CHR(13)+MESSAGE()+CHR(13)+"�������� �� ������ ������ ��������������",48,_vfp.Caption)
		ON ERROR
		QUIT
	
	CASE nError = 2005
		MESSAGEBOX("#2005.������ � ��������� ������"+CHR(13)+MESSAGE()+CHR(13)+"�������� �� ������ ������ ��������������",48,_vfp.Caption)
		ON ERROR
		QUIT
		
	OTHERWISE

	ENDCASE

	LOCAL lcFile as Character, liCnt, nI, lcStr, lcBody, lcCrLf, lcPrg as Character
	lcCrLf = CHR(13)+CHR(10)
	lcFile = ADDBS(SYS(2023)) + "Error_"+DTOC(DATE(),1)+".HTML"

	IF FILE(lcFile)
		DELETE FILE &lcFile
	ENDIF

	lcPrg = []
	STORE 1 TO gnX
	DO WHILE LEN(SYS(16,gnX)) != 0
		lcPrg = lcPrg+CHR(9)+JUSTFNAME(SYS(16,gnX)) + lcCrLf
		STORE gnX+1 TO gnX
	ENDDO

	TEXT TO lcStr NOSHOW TEXTMERGE PRETEXT 4 FLAGS 2
����/����� ������������� ������: <<TTOC(DATETIME())>><br>
���������� ����:<br>
<<lcPrg>><br>
����� ��������� ������: <<cMethod>><br>
������: <<TRANSFORM(nLine)>><br>
����� ������: <<TRANSFORM(nError)>><br>
���������: <<MESSAGE()>><br>
�������������� ���������: <<MESSAGE(1)>><br>
OS:<<OS()>>(<<GETENV("OS")>>)<br>
PROCESSOR_ARCHITECTURE:<<GETENV("PROCESSOR_ARCHITECTURE")>><br>
PROCESSOR_ARCHITEW6432:<<GETENV("PROCESSOR_ARCHITEW6432")>><br>
<<REPLICATE("*",50)>><br>
	ENDTEXT
	IF FILE(lcFile)
		lcStr = lcStr+FILETOSTR(lcFile)
	ENDIF
	STRTOFILE(lcStr,lcFile,0)
TEXT TO lcStr NOSHOW TEXTMERGE PRETEXT 4 FLAGS 2
Date and time current error: <<TTOC(DATETIME())>>
Program stack:
<<lcPrg>>
Method: <<cMethod>>
Line number current error: <<TRANSFORM(nLine)>>
ERROR number: <<TRANSFORM(nError)>>
Messages: <<MESSAGE()>>
----------------------------------------------------------
OS: <<OS()>>(<<GETENV("OS")>>)
PROCESSOR_ARCHITECTURE: <<GETENV("PROCESSOR_ARCHITECTURE")>>
PROCESSOR_ARCHITEW6432: <<GETENV("PROCESSOR_ARCHITEW6432")>>
ENDTEXT
	lcBody = lcStr

	giSmtpPort = 25
	gcSmtpServer = [post]
	gcDefMail = "perminov@ukgres.ru"

	IF FILE("GetAttach.dll")
		DECLARE INTEGER _SendMail IN GetAttach.dll as "SendMail"
		DECLARE _SetParameters IN GetAttach.dll as "SetParameters" STRING lcHost, INTEGER liPort, STRING lcUserName, STRING lcUserPass
		DECLARE _Set_HWND IN GetAttach.dll as "Set_HWND" INTEGER hWND
		DECLARE _SetAuthenticationType IN GetAttach.dll as "SetAuthenticationType" INTEGER liType

		DECLARE _SetAdres IN GetAttach.dll as "SetAdres" STRING lcFrom , STRING lcTo
		DECLARE _SetBody IN GetAttach.dll as "SetBody" STRING lcBody
		DECLARE _SetSubject IN GetAttach.dll as "SetSubject" STRING lcSubject
		DECLARE INTEGER _Attach_Add IN GetAttach.dll as "Attach_Add" STRING lcFile
		DECLARE STRING _GetLastErrorMessage IN GetAttach.dll as "GetLastErrorMessage"
		DECLARE _SetNameFrom IN GetAttach.dll as "SetNameFrom" STRING lcNameFrom
		DECLARE _SetContentType IN GetAttach.dll as "SetContentType" STRING lcContentType

		SetParameters("post",25,"","")
		SetAuthenticationType(0)

		SetContentType("text/text")
		
		SetAdres("error@ukgres.ru", gcDefMail)
		SetNameFrom(SYS(0))
		SetSubject("������ � ���������")
		SetBody(lcBody)
		Set_HWND(_screen.HWnd)
		Attach_Add(lcFile)

		IF SendMail()=1
			MESSAGEBOX("  Error send mail:"+GetLastErrorMessage(),64,This.Name)
		ELSE
	*	                ??" OK"
		ENDIF

		CLEAR DLLS "SendMail"
		CLEAR DLLS "SetParameters"
		CLEAR DLLS "Set_HWND"

		CLEAR DLLS "SetAdres"
		CLEAR DLLS "SetBody"
		CLEAR DLLS "SetSubject"
		CLEAR DLLS "Attach_Add"
		CLEAR DLLS "SetAuthenticationType"
		CLEAR DLLS "GetLastErrorMessage"
		CLEAR DLLS "SetNameFrom"
		CLEAR DLLS "SetContentType"
	ELSE
		post_send_mail(;
			"error@ukgres.ru",;
			gcDefMail,;
			SYS(0)+". ������ � ���������",;
			lcBody,;
			"",;
			lcFile)
	ENDIF

	IF MESSAGEBOX("�������� ������ � ���������"+lcCrLf+"���������� ���������� ������������"+lcCrLf+"����������� � �����"+lcCrLf+lcFile+lcCrLf+lcCrLf+"���������� ������ ���������?",17,_screen.Caption)#1
		ON ERROR
		QUIT
	ELSE
		RETURN .F.
	ENDIF
ENDFUNC

DEFINE CLASS myhandler AS Custom
	Name = "MyErrorHandler"
	
	PROCEDURE MyError
	LPARAMETERS nError, cMethod, nLine
		RETURN ErrorHandler(nError, cMethod, nLine)
	ENDPROC
ENDDEFINE


*
* ����� ��� ����������� � SQL �������
* ��������: ������� ����������� ���������� � ���������� sqlstringconnect()
*
DEFINE CLASS clSQLconnect as Custom
	SQL_ConString = []
	SQL_Comment = [������ �����������]
	SQL_ConnHandler = -1
	SQL_ErrorMessage = []
	SQL_ErrorNumber = 0
	SQL_FileLog = "D:\1\sql.txt"
	CRLF = CHR(13)+CHR(10)
	
	PROCEDURE Init
		LPARAMETERS liH as Integer			&& �������
		
		This.SQL_ConnHandler = liH
		IF This.SQL_ConnHandler <= 0
			MESSAGEBOX("����������� � SQL-������� �� ���������: Handler="+TRANSFORM(This.SQL_ConnHandler)+CHR(13)+"���������� ������ ���������� �� ��������",48,THIS.Name)
			RETURN .F.
		ENDIF
		*STRTOFILE("�����������"+THIS.CRLF, THIS.SQL_FileLog)
	ENDPROC
	
	PROCEDURE Execute
		LPARAMETERS lcSQL as Character, lcAlias as Character
		IF PCOUNT()=1
			lcAlias = []
		ENDIF

		This.SQL_ErrorNumber = 0
		This.SQL_ErrorMessage = []
		
		IF ATC("select ", lcSQL)>0 OR ATC("exec ", lcSQL)>0
			TEXT TO lcSQL1 NOSHOW TEXTMERGE PRETEXT 2
				<<lcSQL>>
			ENDTEXT
		ELSE
			TEXT TO lcSQL1 NOSHOW TEXTMERGE
				BEGIN TRANSACTION
				BEGIN TRY<<SPACE(1)>>
					<<lcSQL>>
					COMMIT TRANSACTION
				END TRY
				BEGIN CATCH
					ROLLBACK TRANSACTION;
					DECLARE @ErrorMessage NVARCHAR(4000);
					DECLARE @ErrorSeverity INT;
					DECLARE @ErrorState INT;
					SELECT @ErrorMessage = ERROR_MESSAGE();
					SELECT @ErrorSeverity = ERROR_SEVERITY();
					SELECT @ErrorState = ERROR_STATE();
					RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);
				END CATCH
			ENDTEXT
		ENDIF
		*This.SQL_Log.Puts("���������� SQL-�������>>")
		*This.SQL_Log.Puts(lcSQL1)
		LOCAL liResult as Integer
		IF TYPE("lcAlias")="C" AND USED(lcAlias) AND CURSORGETPROP("Buffering", lcAlias) > 3
			TABLEREVERT(.T., lcAlias)
		ENDIF
		liResult = SQLEXEC(This.SQL_ConnHandler, lcSQL1, lcAlias)
		
		*STRTOFILE("����������:"+THIS.CRLF, THIS.SQL_FileLog,1)
		*STRTOFILE(lcSQL1+THIS.CRLF, THIS.SQL_FileLog,1)
		*STRTOFILE("���������: "+TRANSFORM(liResult)+THIS.CRLF, THIS.SQL_FileLog,1)
		*STRTOFILE("Alias "+lcAlias+THIS.CRLF, THIS.SQL_FileLog,1)

		DO CASE
		CASE liResult = -1
			DIMENSION laErr[1]
			AERROR(laErr)
			MESSAGEBOX(laErr[2]+CHR(13)+lcSQL1,64,THIS.Name)
			This.SQL_ErrorNumber = laErr[1]
			This.SQL_ErrorMessage = laErr[2]
			*This.SQL_Log.Puts("Error:"+laErr[2])

		CASE liResult = 0
		CASE liResult = 1
			*This.SQL_Log.Puts("Ok")
		ENDCASE
		RETURN (liResult)
	ENDPROC
ENDDEFINE

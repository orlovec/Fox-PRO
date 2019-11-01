* PROCEDUR.PRG
* Процедурный файл

* Создание папки с обработкой ошибки
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



* Установка времени для файла
*
PROCEDURE new_filetime(lFileName, lFileTime)  
 * lFileName путь к файлу, формат chr  
 * lFileTime новое время, формат datetime   
    
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
  
*!*	* Остановить проигрывание
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

*!*	* Начать проигрывание
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

*!*	* Остановить проигрывание
*!*	mciExecute("close pinmidi.wav")
CLEAR DLLS "mciExecute"
RETURN (.T.)
ENDFUNC

*
* Делаем строку из диапазона чисел: 2013,2014,2015
* Параметры: liNum_Start-начальное число, liNum_End-конечное число, lbOrder-порядок: .T.-с возрастанием (2013,2014,2015), .F.-с убыванием (2015,2014,2013)
* Дата: 08-08-2012
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
	_s = "Понедельник"
CASE _n = 2
	_s = "Вторник"
CASE _n = 3
	_s = "Среда"
CASE _n = 4
	_s = "Четверг"
CASE _n = 5
	_s = "Пятница"
CASE _n = 6
	_s = "Суббота"
CASE _n = 7
	_s = "Воскресенье"
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
cM[1] = "января"
cM[2] = "февраля"
cM[3] = "марта"
cM[4] = "апреля"
cM[5] = "мая"
cM[6] = "июня"
cM[7] = "июля"
cM[8] = "августа"
cM[9] = "сентября"
cM[10] = "октября"
cM[11] = "ноября"
cM[12] = "декабря"
cStr = lzerro(DAY(dT),2)+' '+cM[MONTH(dT)]+' '+STR(YEAR(dT),4)+' г.'
RETURN (cStr)
ENDFUNC

* Преобразование названия месяца в номер месяца
FUNCTION MonthFromStr(lcS)
LOCAL _s
_s = 0
DO CASE
CASE ATC("январ",lcS)>0
	_s = 1
CASE ATC("феврал",lcS)>0
	_s = 2
CASE ATC("март",lcS)>0
	_s = 3
CASE ATC("апрель",lcS)>0
	_s = 4
CASE (ATC("май",lcS)>0 OR ATC("мая",lcS)>0)
	_s = 5
CASE ATC("июнь",lcS)>0
	_s = 6
CASE ATC("июль",lcS)>0
	_s = 7
CASE ATC("август",lcS)>0
	_s = 8
CASE ATC("сентябрь",lcS)>0
	_s = 9
CASE ATC("октябр",lcS)>0
	_s = 10
CASE ATC("ноябр",lcS)>0
	_s = 11
CASE ATC("декабр",lcS)>0
	_s = 12
ENDCASE
RETURN (_s)
ENDFUNC

* Преобразование номера месяца в название с использованием склонения
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
cM[1,1] = "январь"
cM[1,2] = "января"

cM[2,1] = "февраль"
cM[2,2] = "февраля"

cM[3,1] = "март"
cM[3,2] = "марта"

cM[4,1] = "апрель"
cM[4,2] = "апреля"

cM[5,1] = "май"
cM[5,2] = "мая"

cM[6,1] = "июнь"
cM[6,2] = "июня"

cM[7,1] = "июль"
cM[7,2] = "июля"

cM[8,1] = "август"
cM[8,2] = "августа"

cM[9,1] = "сентябрь"
cM[9,2] = "сентября"

cM[10,1] = "октябрь"
cM[10,2] = "октября"

cM[11,1] = "ноябрь"
cM[11,2] = "ноября"

cM[12,1] = "декабрь"
cM[12,2] = "декабря"
RETURN (cM[nM,nPos])
ENDFUNC


FUNCTION RusMonth(nM)
IF PCOUNT() = 0
	nM = MONTH(DATE())
ENDIF
LOCAL cM
DIMENSION cM[12]
cM[1] = "Январь"
cM[2] = "Февраль"
cM[3] = "Март"
cM[4] = "Апрель"
cM[5] = "Май"
cM[6] = "Июнь"
cM[7] = "Июль"
cM[8] = "Август"
cM[9] = "Сентябрь"
cM[10] = "Октябрь"
cM[11] = "Ноябрь"
cM[12] = "Декабрь"
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
* Принимает: Путь
* Если его нет то автоматически его создает
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


* Функция FILEDBF
* Параметры: Год, Месяц, Название
* Генерируется название типа 9704DBF
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
* Параметры: целое число, длинна
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
* Установливает указатель на необходимую запись
* Параметры: <Номер записи>,[<Псевдоним>]
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
* Проверка и запись в файл DBF русской кодовой страницы
* для DOS - 866 (0x65, 101)
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

* -- Возвращает название месяца
Function MonthToC
parameters _m
if pcount() = 0
	_m = month(date())
endif
local _name
do case
case _m = 1
	_name = "январь"
case _m = 2
	_name = "февраль"
case _m = 3
	_name = "март"
case _m = 4
	_name = "апрель"
case _m = 5
	_name = "май"
case _m = 6
	_name = "июнь"
case _m = 7
	_name = "июль"
case _m = 8
	_name = "август"
case _m = 9
	_name = "сентябрь"
case _m = 10
	_name = "октябрь"
case _m = 11
	_name = "ноябрь"
case _m = 12
	_name = "декабрь"
endcase
return _name


* FUNCTION PROPIS_CIEL 
* Сумма цыфрами PROPIS_CIEL(152.35) = 152 руб. 35 коп.
* Возвращает строку символов
*
Function Propis_Ciel
lparameters lnZn
if type("lnZn") # "N"
	return ""
endif
local lnLeft, lnRight
lnLeft = int(lnZn)
lnRight = int((lnZn - lnLeft) * 100)
return alltrim(str(lnLeft)) + " руб. " + transform(lnRight, "@RL 99 коп.")
endfunc

* PROPIS FUNCTION
* Сумма прописью
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
        _2=_2+'ОДИН МИЛЛИАРД '
case _0=2
        _2=_2+'ДВА МИЛЛИАРДА '
case _0=3
        _2=_2+'МИЛЛИАРДА '
case _0=4
        _2=_2+'МИЛЛИАРДОВ '
endcase
_2=_2+SIM(substr(_1,4,3))
do case
case _0=1
        _2=_2+'ОДИН МИЛЛИОН '
case _0=2
        _2=_2+'ДВА МИЛЛИОНА '
case _0=3
        _2=_2+'МИЛЛИОНА '
case _0=4
        _2=_2+'МИЛЛИОНОВ '
endcase
_2=_2+SIM(substr(_1,7,3))
do case
case _0=1
        _2=_2+'ОДНА ТЫСЯЧА '
case _0=2
        _2=_2+'ДВЕ ТЫСЯЧИ '
case _0=3
        _2=_2+'ТЫСЯЧИ '
case _0=4
        _2=_2+'ТЫСЯЧ '
endcase
_2=_2+SIM(substr(_1,10,3))
do case
case _0=0.and.empty(_2)
        _2='НУЛЬ РУБЛЕЙ '
case _0=1
        _2=_2+'ОДИН РУБЛЬ '
case _0=2
        _2=_2+'ДВА РУБЛЯ '
case _0=3
        _2=_2+"РУБЛЯ "
case _0=4
        _2=_2+"РУБЛЕЙ "
otherwise
        _2=_2+"РУБЛЕЙ "
endcase
_0 = substr(_1,15,1)
do case
case _0 = "1"
        _0 = " КОПЕЙКА"
case _0 >= "2" and _0 <= "4"
        _0 = " КОПЕЙКИ"
otherwise
        _0 = ' КОПЕЕК'
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
        SIM='СТО '
case CIF=2
        SIM='ДВЕСТИ '
case CIF=3
        SIM='ТРИСТА '
case CIF=4
        SIM='ЧЕТЫРЕСТА '
case CIF=5
        SIM='ПЯТЬСОТ '
case CIF=6
        SIM='ШЕСТЬСОТ '
case CIF=7
        SIM='СЕМЬСОТ '
case CIF=8
        SIM='ВОСЕМЬСОТ '
case CIF=9
        SIM='ДЕВЯТЬСОТ '
endcase
CIF=val(substr(SIM_,2,1))
do case
case CIF=2
        SIM=SIM+'ДВАДЦАТЬ '
case CIF=3
        SIM=SIM+'ТРИДЦАТЬ '
case CIF=4
        SIM=SIM+'СОРОК '
case CIF=5
        SIM=SIM+'ПЯТЬДЕСЯТ '
case CIF=6
        SIM=SIM+'ШЕСТЬДЕСЯТ '
case CIF=7
        SIM=SIM+'СЕМЬДЕСЯТ '
case CIF=8
        SIM=SIM+'ВОСЕМЬДЕСЯТ '
case CIF=9
        SIM=SIM+'ДЕВЯНОСТО '
endcase
_0=4
if CIF=1
        CIF=val(substr(SIM_,3,1))
        do case
        case CIF=0
                SIM=SIM+'ДЕСЯТЬ '
        case CIF=1
                SIM=SIM+'ОДИННАДЦАТЬ '
        case CIF=2
                SIM=SIM+'ДВЕНАДЦАТЬ '
        case CIF=3
                SIM=SIM+'ТРИНАДЦАТЬ '
        case CIF=4
                SIM=SIM+'ЧЕТЫРНАДЦАТЬ '
        case CIF=5
                SIM=SIM+'ПЯТНАДЦАТЬ '
        case CIF=6
                SIM=SIM+'ШЕСТНАДЦАТЬ '
        case CIF=7
                SIM=SIM+'СЕМНАДЦАТЬ '
        case CIF=8
                SIM=SIM+'ВОСЕМНАДЦАТЬ '
        case CIF=9
                SIM=SIM+'ДЕВЯТНАДЦАТЬ '
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
                SIM=SIM+'ТРИ '
        case CIF=4
                _0=3
                SIM=SIM+'ЧЕТЫРЕ '
        case CIF=5
                SIM=SIM+'ПЯТЬ '
        case CIF=6
                SIM=SIM+'ШЕСТЬ '
        case CIF=7
                SIM=SIM+'СЕМЬ '
        case CIF=8
                SIM=SIM+'ВОСЕМЬ '
        case CIF=9
                SIM=SIM+'ДЕВЯТЬ '
        endcase
endif
return SIM

* Запись координат окна и размеров
* WSave(oFormRef)
* oFormRef			Сласс окна (Thisform)
* Пример: WSave(thisform)
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

* Чтение координат окна и размеров, для записи используйте функцию WSave(iUserID, oFormRef)
* WGet(oFormRef)
* oFormRef		Сласс окна (This)
* Пример: WGet(thisform)
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

* Функция обратная Period_Dnei()
* Возвращает дату: DD.MM.YYYY
FUNCTION Period_Dnei_Dt
LPARAMETERS lcStr as Character
PRIVATE _Dd, _Mm, _Yy
_Dd = 1
DO CASE
CASE ATC("январь",lcStr)>0 OR ATC("января",lcStr)>0 OR ATC("январ",lcStr)>0
	_Mm = 1
CASE ATC("февраль",lcStr)>0 OR ATC("февраля",lcStr)>0 OR ATC("феврал",lcStr)>0
	_Mm = 2
CASE ATC("март",lcStr)>0 OR ATC("марта",lcStr)>0
	_Mm = 3
CASE ATC("апрель",lcStr)>0 OR ATC("апреля",lcStr)>0 OR ATC("апрел",lcStr)>0
	_Mm = 4
CASE ATC("май",lcStr)>0 OR ATC("мая",lcStr)>0
	_Mm = 5
CASE ATC("июнь",lcStr)>0 OR ATC("июня",lcStr)>0
	_Mm = 6
CASE ATC("июль",lcStr)>0 OR ATC("июля",lcStr)>0
	_Mm = 7
CASE ATC("август",lcStr)>0 OR ATC("августа",lcStr)>0
	_Mm = 8
CASE ATC("сентябрь",lcStr)>0 OR ATC("сентября",lcStr)>0 OR ATC("сентяб",lcStr)>0
	_Mm = 9
CASE ATC("октябрь",lcStr)>0 OR ATC("октября",lcStr)>0 OR ATC("октяб",lcStr)>0
	_Mm = 10
CASE ATC("ноябрь",lcStr)>0 OR ATC("ноября",lcStr)>0 OR ATC("ноябр",lcStr)>0
	_Mm = 11
CASE ATC("декабрь",lcStr)>0 OR ATC("декабря",lcStr)>0 OR ATC("декабр",lcStr)>0
	_Mm = 12
ENDCASE
_Yy = VAL(SUBSTR(lcStr,AT(SPACE(1),lcStr,5)+1,4))
RETURN (DATE(_Yy,_Mm, 1))
ENDFUNC

* Определение периода дней в месяце, например: с 01 по 30 апреля 2013 года
* Period_Dnei(ldDt)
* ldDt - дата, из даты берется только месяц и год
FUNCTION Period_Dnei
LPARAMETERS _dp
LOCAL _m, _y, _lday
_m = MONTH(_dp)
_y = YEAR(_dp)
_lday = LDay(_dp)

RETURN "с 01 по "+TRANSFORM(_lday,"@L 99")+" "+RusMonth_EX(_m,2)+" "+TRANSFORM(_y, "@L 9999")+" года"
ENDFUNC

* Определение кол-ва дней в месяце
* LDay(ldDt)
* ldDt - дата
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

* Определение высокосный год или нет
* Vis(dDt)
* dDt дата
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
* Установка шрифта для объекта
* Параметры:
* 1 - объект
* 2 - строка шрифта смотри getfont()
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
* Чтение параметров из файла access.ass
* Параметры
* 1 - Секция, например "[Test]"
* 2 - Параметр, например "FontName"
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
					* Началась другая секция ?
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
* Запись параметров в файла access.ass
* Параметры
* 1 - Секция, например "[Test]"
* 2 - Параметр, например "FontName"
* 3 - Значение
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
			* Нашли искомую секцию
			if _myAss[i] == _Section
				i = i + 1
				for j=i to alen(_myAss)
					* Началась другая секция ?
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


*!*	*** функция GUID() - возвращает 8 байтовое уникальное значение 
*!*	*** GUID - функция от времени (с быстр. не хуже 0,1 мс), сессии и кода ПК
*!*	*** также доступны: DatFromGuid() - восстановление даты из GUID
*!*	*** необходимо подключить WSUBD.FLL (SET LIBRARY TO ...)
*!*	*** Кубанский госуниверситет, (с) Дм.Баянов, 1998


*!*	****************************************************************
*!*	*** Инициализация генератора GUID - вызывать в начале приложения
*!*	Function IniGenGuid
*!*	Public hid2_guid, UniString_, nd96s, nd96d && счетчики для идентификатора GUID
*!*	 nd96s=0
*!*	 nd96d=0
*!*	 *** Загрузка массива символов для уникального кода
*!*	 UniString_=GetUniString() && сохраним для быстрого извлечения
*!*	 *** инициализация 
*!*	 cBuf=Sys(2015) && подмешаем в аппаратный код код сессии
*!*	  ** GetVolInfo - получить информацию о HARD (hdd)
*!*	  ** GetICcrc16 - получить контр. сумму текста
*!*	  ** NumToB     - упаковать число в строку
*!*	 ** Hid2_Guid - 2 байта контр. суммы ПК и сессии запуска
*!*	 Hid2_Guid=PADR(NumtoB(GetICrc16(GetVolInfo()+cBuf)),2,'*')&& low функции - в файле wsubd.fll
*!*	 nd96s=Tic96() && последнее значение тиков
*!*	RETURN

*!*	*************************************************************************
*!*	* Основная функция - Получить уникальный идентификатор GUID = Hard+Time
*!*	Function GUID
*!*	* До 2010 года
*!*	* 6b временной счетчик + 2b hard-конфигурации+сессии
*!*	Return PADR(NumtoB(Tic96n())+Hid2_guid,8,'!')


*!*	************************************
*!*	** Восстановить десятичное число из GUID (тики + 2разряда секунды + 2 разряда счетчик доп)
*!*	Function NumFromGuid(cB)
*!*	RETURN BTONum(SUBSTR(cb,1,6))


*!*	************************************
*!*	** Восстановить дату из Guid
*!*	Function DatFromGuid(cB)
*!*	Local nTic, dDate
*!*	nTic=BTONum(SUBSTR(cb,1,6))
*!*	nTic=nTic/100 && усечем 2 доп.разряда счетчика
*!*	dDate=INT(nTic/8640000) && дни
*!*	dDate=CTOD(SYS(10,dDate+2450800))
*!*	RETURN dDate


*!*	*********************************************************
*!*	*** Построение массива символов для уникального кода
*!*	Function GetUniString
*!*	Local cUniS,ni
*!*	    && cUnis - это набор знаков для отображения чисел с основанием LEN(cInis)
*!*	    && выкинули символы: '"[]&   (39,34,91,93)
*!*	    *** набираем строку знаков
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
*!*	* Получить число тиков (10 мс - точнее не может NT) прошедшее с 1998 года 
*!*	* 6 байт - хватает на 20 лет
*!*	* 2451545 - число дней, прошедших к 01-01-2000
*!*	Function Tic96
*!*	return int(val(sys(11,date()-2451545))*8640000+seconds()*100)


*!*	**********************************************************
*!*	* Получить уникальное число - тики Tic96 + счетчик-дополнение
*!*	* В пределах одного тика Tic96 может вызваться несколько раз
*!*	* поэтому вводим дополнительно increment, old - in nd96
*!*	Function Tic96N
*!*	Local nNew,nNewS
*!*	 nNewS=Tic96()
*!*	  IF nNewS==nd96s && все еще старый тик
*!*	    IF nd96d==99 && уже нет запаса счетчика - тянем новый тик
*!*	       DO WHILE nNewS==nd96s
*!*	         nNewS=Tic96()
*!*	       ENDDO
*!*	      nd96d=0
*!*	    ELSE
*!*	      nd96d=nd96d+1 && Запас есть - увеличиваем счетчик
*!*	    ENDIF  
*!*	  ELSE && новый тик
*!*	    nd96d=0  
*!*	  ENDIF
*!*	       nNew=nNewS*100+nd96d  && Получим новый Id
*!*	       nd96s=nNewS
*!*	return (nNew)


*!*	************************************************************
*!*	* Преобразование числа nNum в строку символов
*!*	* это фактически упаковка числа в строку символов
*!*	* строка используемых знаков находится в переменной UniString_
*!*	* (c) Dm.Bayanov 1997 , 1998
*!*	* (ф-я VFP BinToC() вылетает при числе разрядов > 9)
*!*	* обратное преобразование - BtoNum
*!*	*** Пример выбора систем представлени
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
*!*	*** Основание итогового числа - длина массива используемых знаков UniString
*!*	nOsn=Len(cUniStr)&&-1 && диапазон визуальных символов для строки
*!*	*** Вычислим количество разрядов в итоговом числе (ex.999>123)
*!*	DO WHILE .T.
*!*	    nOst=nOst+(nOsn-1)*nOsn**(nLen-1)
*!*	 IF nOst >= nNum
*!*	   EXIT
*!*	 ENDIF
*!*	  nLen=nLen+1
*!*	ENDDO
*!*	   FOR ni=1 TO nLen
*!*	     nOst=nNum-nTale
*!*	      * вес ni разряда - каждый разряд может принимать значение
*!*	      *  от старщего до младшего элемента UniString
*!*	     nVes=INT(nOst/nOsn**(nLen-ni))
*!*	      * текущий остаток от границы текущего разряда
*!*	     nTale=nTale+nVes*nOsn**(nLen-ni)
*!*	         *  по весу разряда извлекаем символ из строки уникальных знаков UniStr
*!*	         * nVes+1 для получения знака по nVes==0
*!*	        cBuf=cBuf+SUBSTR(cUniStr,nVes+1,1)
*!*	   NEXT
*!*	RETURN cBuf


*!*	*********************************************
*!*	** Обратное преобразование в десятичное число
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
* Запись в файл INI
*
* Параметры:
* tcFileName		- файл с расширением .INI, если просто файл без пути, то пишется в каталог %OS%
* tcSection			- секция
* tcEntry				- параметр
* tcValue				- значение параметра
* Возвращает:
* .t. - OK иначе .f.
*
FUNCTION WriteFileIni(tcFileName,tcSection,tcEntry,tcValue)
DECLARE INTEGER WritePrivateProfileString ;	
	IN WIN32API ;
	STRING cSection,STRING cEntry,STRING cEntry,;	
	STRING cFileName
RETURN IIF(WritePrivateProfileString(tcSection,tcEntry,tcValue,tcFileName)=1, .T., .F.)

ENDFUNC

*
* Чтение из файла INI
*
* Парметры:
* tcFileName			- файл с расширением .INI, если просто файл без пути, то берётся из каталога %OS%
* tcSection				- секция
* tcEntry				- параметр
* Возвращает:
* занчение параметра tcEntry или .NULL.
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
* Переключение на русский в автоматическом режиме
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
* Переключение на US раскладку клавиатуры
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
* Текущий код раскладки клавиатуры
* 409 - англ
* 419 - рус.
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
* Создаёт курсор со списком таблиц которые находятся в открытой БД
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


*! Подключение сетевого диска
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
* <STR_1> строку символов
* ?RIGHT_P('ПЕРМИНОВ ИГОРЬ ЭНГЕЛЬСОВИЧ')
* ПЕРМИНОВ И.Э.
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
* Проверка повторного запуска программы
*
* IF FirstInstance("Narkota") == .T.
* 	DO FORM SYS(5)+SYS(2003)+'\Narkota'
* 	READ EVENTS
* ELSE
* 	MessageBox("Вы уже запустились один раз! Видимо достаточно.")
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
	caption = 'Сообщение'
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
* Функция преобразования фамилии, имя, отчества
* в форму нужного падежа
*
* Аргументы:
* mFamily - фамилия
* mName - имя
* mSoName - отчество
* Pol - пол (допустимые значения "м", "ж")
* Padeg - падеж (допустимые значения 1..6)
* && 1-именительный
* && 2-родительный
* && 3-дательный
* && 4-винительный
* && 5-творительный
* && 6-предложный
*********************************************************

Function FIO
Parameter mFamily, mName, mSoName, Pol, Padeg

mFamily=AllTrim(mFamily)
mName=AllTrim(mName)
mSoName=AllTrim(mSoName)

If Between(Padeg, 1, 6)
	Pol=Lower(Pol)
	Do Case
	Case Pol="м"
		mFio=Fam_m(mFamily, padeg) && преобразуем фамилию
		mFio=mFio + " " + Name_m(mName, padeg) && преобразуем имя
		mFio=mFio + " " + SoName_m(mSoName, padeg) && преобразуем отчество
	Case Pol="ж"
		mFio=Fam_w(mFamily, padeg) && преобразуем фамилию
		mFio=mFio + " " + Name_w(mName, padeg) && преобразуем имя
		mFio=mFio + " " + SoName_w(mSoName, padeg) && преобразуем отчество
	OtherWise && если не указан пол, выдаем сообщение об ощибке
		mFio="Hеверно указан пол !"
	EndCase
ELSE
	mFio="Hедопустимое значение падежа !"
EndIf
Return mFio




********************************************
* Преобразование мужской фамилии
********************************************
Function Fam_m
Parameter Fam, Pad && параметры фамилия в именит. падеже
&& и требуемый падеж
&& 1-именительный
&& 2-родительный
&& 3-дательный
&& 4-винительный
&& 5-творительный
&& 6-предложный

Declare Oki(6) && определяем массив окончаний
Store "" to Oki

Fam_m=""
Fam=Proper(Fam)
Len=Len(Fam) && определяем количество букв в фамилии

Ok=Substr(Fam, Len-2, 3) && берем последние 3 буквы фамилии
If Ok="кий" .OR. Ok="ний"; && если окончание такое

Fam=Substr(Fam, 1, len-2) && то формируем фамилию
Do Case
Case Ok="кий"
	Oki[1]="ий"
	Oki[2]="ого"
	Oki[3]="ому"
	Oki[4]="ого"
	Oki[5]="им"
	Oki[6]="ом"
Case Ok="кий"
	Oki[1]="ий"
	Oki[2]="его"
	Oki[3]="ему"
	Oki[4]="его"
	Oki[5]="им"
	Oki[6]="ем"
EndCase
Fam_m=Fam+Oki[Pad] && добавляем к фамилии окончание
Return Fam_m && возвращем результат
EndIf

***** перебираем другие окончания
Ok=Substr(Fam, Len-1, 2) && берем последние две буквы
If Ok="ий"
	Fam=Substr(Fam, 1, Len-1) && формируем фамилию
	Oki[1]="й"
	Oki[2]="я"
	Oki[3]="ю"
	Oki[4]="я"
	Oki[5]="ём"
	Oki[6]="е"
	Fam_m=Fam+Oki[Pad] && добавляем к фамилии окончание
	Return Fam_m && возвращем результат
EndIf

Do Case
Case Ok="ын" .OR. Ok="ин" .OR. Ok="ев" .OR. Ok="ёв" .OR. Ok="ов"
	Oki[1]=""
	Oki[2]="а"
	Oki[3]="у"
	Oki[4]="а"
	Oki[5]="ым"
	Oki[6]="е"
Case Ok="ян" .OR. Ok="ан" .OR. Ok="он" .OR. Ok="ук" .OR. Ok="юк" .OR. Ok="яр"
	Oki[1]=""
	Oki[2]="а"
	Oki[3]="у"
	Oki[4]="а"
	Oki[5]="ом"
	Oki[6]="е"
Case Ok="ок"
	If Pad>1
		Fam=Substr(Fam, 1 , len-2)+"к"
	EndIf
	Oki[1]=""
	Oki[2]="а"
	Oki[3]="у"
	Oki[4]="а"
	Oki[5]="ом"
	Oki[6]="е"
Case Ok="ый" .OR. Ok="ой"
	Fam=Substr(Fam, 1, Len-2)
	Oki[1]=""
	Oki[2]="ого"
	Oki[3]="ому"
	Oki[4]="ого"
	Oki[5]="ым"
	Oki[6]="ом"
Case Ok="ич"
	Oki[1]=""
	Oki[2]="а"
	Oki[3]="у"
	Oki[4]="а"
	Oki[5]="ем"
	Oki[6]="е"
EndCase
Fam_m=Fam+Oki[Pad]
Return Fam_m



********************************************
* Преобразование женской фамилии
********************************************
Function Fam_w

Parameter Fam, Pad && параметры фамилия в имен. падеже
&& и требуемый падеж
&& 1-именительный
&& 2-родительный
&& 3-дательный
&& 4-винительный
&& 5-творительный
&& 6-предложный

Declare Oki(6) && определяем массив окончаний
Store "" to Oki

Fam_m=""
Fam=Proper(Fam)
Len=Len(Fam) && определяем количество букв в фамилии

Ok=Substr(Fam, Len-1, 2) && берем посление 3 буквы фамилии
If Ok="ая" .OR. Ok="яя"; && если окончание такое

Fam=Substr(Fam, 1, len-2) && то формируем фамилию
Do Case
Case Ok="ая"
Oki[1]="ая"
Oki[2]="ой"
Oki[3]="ой"
Oki[4]="ую"
Oki[5]="ой"
Oki[6]="ой"
Case Ok="яя"
Oki[1]="яя"
Oki[2]="ей"
Oki[3]="ей"
Oki[4]="юю"
Oki[5]="ей"
Oki[6]="ей"
EndCase
Fam_m=Fam+Oki[Pad] && добавляем к фамилии окончание
Return Fam_m && возвращем результат
EndIf


***** перебираем другие окончания

Ok=Substr(Fam, Len-2, 3) && берем последние две буквы

If Ok="ова" .OR. Ok="ева" .OR. Ok="ёва" .OR. Ok="ина"
Fam=Substr(Fam, 1, Len-1) && формируем фамилию
Oki[1]=Ok
Oki[2]="ой"
Oki[3]="ой"
Oki[4]="у"
Oki[5]="ой"
Oki[6]="ой"
EndIf

Fam_m=Fam+Oki[Pad]
Return Fam_m



**************************************************
* Функция преобразования мужского имени
*
*
**************************************************

Function Name_m
Parameter Name, Pad && параметры: имя в имен. падеже
&& и требуемый падеж
&& 1-именительный
&& 2-родительный
&& 3-дательный
&& 4-винительный
&& 5-творительный
&& 6-предложный

Declare Oki(6) && определяем массив окончаний
Store "" to Oki

Name=Proper(Name)
Len=Len(Name)

Ok=Substr(Name, Len, 1) && смотрим окончание
Do Case
Case Ok="й"
Oki[1]=""
Oki[2]="я"
Oki[3]="ю"
Oki[4]="я"
Oki[5]="ем"
Oki[6]="е"
Name=Substr(Name, 1, Len-1)
OtherWise
Oki[1]=""
Oki[2]="а"
Oki[3]="у"
Oki[4]="а"
Oki[5]="ом"
Oki[6]="е"
EndCase
Name_m=Name+Oki[Pad]
Return Name_m



**************************************************
* Функция преобразования женского имени
*
*
**************************************************

Function Name_w
Parameter Name, Pad && параметры: имя в имен. падеже
&& и требуемый падеж
&& 1-именительный
&& 2-родительный
&& 3-дательный
&& 4-винительный
&& 5-творительный
&& 6-предложный

Declare Oki(6) && определяем массив окончаний
Store "" to Oki

Name=Proper(Name)
Len=Len(Name)

Ok=Substr(Name, Len, 1) && смотрим окончание
Do Case
Case Ok="а"
Oki[1]=""
Oki[2]="ы"
Oki[3]="е"
Oki[4]="у"
Oki[5]="ой"
Oki[6]="е"
Name=Substr(Name, 1, Len-1)
Case Ok="я"
Oki[1]=""
Oki[2]="и"
Oki[3]="е"
Oki[4]="ю"
Oki[5]="ей"
Oki[6]="и"
Name=Substr(Name, 1, Len-1)
EndCase
Name_m=Name+Oki[Pad]
Return Name_m




**************************************************
* Функция преобразования мужского отчества
*
*
**************************************************

Function SoName_m
Parameter SoName, Pad && параметры: имя в имен. падеже
&& и требуемый падеж
&& 1-именительный
&& 2-родительный
&& 3-дательный
&& 4-винительный
&& 5-творительный
&& 6-предложный

Declare Oki(6) && определяем массив окончаний
Store "" to Oki

SoName=Proper(SoName)
Len=Len(SoName)

Oki[1]=""
Oki[2]="а"
Oki[3]="у"
Oki[4]="а"
Oki[5]="ем"
Oki[6]="е"

SoName_m=SoName+Oki[Pad]
Return SoName_m




**************************************************
* Функция преобразования женского отчества
*
*
**************************************************

Function SoName_w
Parameter SoName, Pad && параметры: имя в имен. падеже
&& и требуемый падеж
&& 1-именительный
&& 2-родительный
&& 3-дательный
&& 4-винительный
&& 5-творительный
&& 6-предложный

Declare Oki(6) && определяем массив окончаний
Store "" to Oki

SoName=Proper(SoName)
Len=Len(SoName)

Ok=Substr(SoName, Len, 1) && смотрим окончание
If Ok="а"
Oki[1]="а"
Oki[2]="ы"
Oki[3]="е"
Oki[4]="у"
Oki[5]="ой"
Oki[6]="е"
SoName=Substr(SoName, 1, Len-1)
EndIf
SoName_w=SoName+Oki[Pad]
Return SoName_w

*
* Преобразование русских букв в транслит. Вариант 1
*
FUNCTION Rus2Translit
LPARAMETERS lpValue  
  CREATE CURSOR tBukv (cKir C(2), cLat C(2))  
  INSERT INTO tBukv (cKir, cLat) VALUES ("й","y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ц","c")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("у","u")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("к","k")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("е","e")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("н","n")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("г","g")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ш","sh")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("щ","sh")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("з","z")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("х","h")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ф","f")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ы","i")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("в","v")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("а","a")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("п","p")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("р","r")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("о","o")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("л","l")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("д","d")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ж","j")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("э","e")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("я","ya")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ч","ch")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("с","s")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("м","m")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("и","i")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("т","t")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ь","y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("б","b")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ю","yu")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("ъ","")  
 *********************************************************  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Й","Y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ц","C")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("У","U")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("К","K")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Е","E")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Н","N")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Г","G")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ш","SH")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Щ","SH")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("З","Z")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Х","H")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ф","F")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ы","I")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("В","V")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("А","A")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("П","P")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Р","R")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("О","O")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Л","L")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Д","D")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ж","J")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Э","E")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Я","YA")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ч","CH")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("С","S")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("М","M")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("И","I")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Т","T")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ь","Y")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Б","B")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ю","YU")  
  INSERT INTO tBukv (cKir, cLat) VALUES ("Ъ","")  
 *********************************************************  
  SELECT "tBukv"  
  SCAN  
    lpValue=STRTRAN(lpValue, ALLTRIM(tBukv.cKir), ALLTRIM(tBukv.cLat))  
  endSCAN  
  USE IN "tBukv"  
  RETURN lpValue
  
*
* Преобразование русских букв в транслит. Вариант 2
*
FUNCTION translit
  LPARAMETERS _cString  
  _cString=UPPER(_cString)  
  LOCAL _cRetStr,_a,_b,a1,a2,a3,a4,b1,b2,b3,b4  
  a1="Ш"  
  a2="Щ"  
  a3="Я"  
  a4="Ч"  
    
  b1="SH"  
  b2="CH"  
  b3="YA"  
  b4="CH"  
    
  _a="ЙЦУКЕНГЗХЇФІВАПРОЛДЖЄСМИТЬБЮ№"  
  _b="YCUKENGZHIFIVAPROLDJESMIT'BUN"  
    
  _cRetStr=CHRTRAN(_cString,_a,_b)  
  _cRetStr=STRTRAN(_cRetStr,a1,b1)  
  _cRetStr=STRTRAN(_cRetStr,a2,b2)  
  _cRetStr=STRTRAN(_cRetStr,a3,b3)  
  _cRetStr=STRTRAN(_cRetStr,a4,b4)  
  RETURN _cRetStr  

*
* Net Time или функция для получения времени с сервера в сети на основе API NetRemoteTOD
* 08-10-2004
* http://www.foxclub.ru/sol/index.php?act=view&id=425
*
FUNCTION NetTime(m.cServerName)
*Имя сервера обязательно с символов \\, например \\server

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
IF NetRemoteTOD(m.lcServerName,@m.lnBufferPointer)=0 && Все хорошо!
m.nHour = GetMemoryDWORD(lnBufferPointer+8) &&часы
m.nMin = GetMemoryDWORD(lnBufferPointer+12) &&минуты
m.nSec = GetMemoryDWORD(lnBufferPointer+16) &&секунды
m.nmDiv = GetMemoryDWORD(lnBufferPointer+24) &&смещение
m.nDay = GetMemoryDWORD(lnBufferPointer+32) &&день
m.nMon = GetMemoryDWORD(lnBufferPointer+36) &&месяц
m.nYear = GetMemoryDWORD(lnBufferPointer+40) &&год
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
 * Форматирование даты с переводом слов на указанный язык  
 * Параметры:  
 * 1 - Форматируемая дата, допустимо передавать дату-время,  
 *  и просто номер месяца.  
 *  По умолчанию используется текущая системная дата.  
 * 2 - LCID (1033 - английский, 1049 - русский, 1059 - белорусский),  
 *  также допустимо указывать "S" и "D" для соответственно  
 *  системной и пользовательской локали.  
 *  По умолчанию используется пользовательская локаль.  
 * 3 - Формат, как описано в MSDN. Наиболее общие форматы это:  
 *  MMMM - полное название месяца (используется по умолчанию)  
 *  MMM - сокращённое название месяца  
 *  dddd - полное название дня недели  
 *  ddd - сокращённое название дня недели  
 *  
 * Используется в:  
 *  
 * Вызывает: SWord()  
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
  lnShow=IIF(TYPE("lnShow")="N",lnShow,0) && 0 - выполнять тихо  
  cParameters=IIF(TYPE("cParameters")="C",alltrim(cParameters),"")+chr(0) && параметры (ключи)  
  nWinHandle=IIF(TYPE("nWinHandle")="N",nWinHandle,0) && см. api FindWindow  
    
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
*        =MESSAGEBOX ("Создание архива завершено")

*
* Создаем новую сессию данных
*
FUNCTION MakeNewDataSession
	oSes=CREATEOBJECT("session")  
	SET DATASESSION TO oSes.DataSession  
	RETURN (oSes)
ENDFUNC

* Функция обработки события
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
	* ПРОЦЕСС
	DECLARE integer PeekMessage IN WIN32API AS WinAPI_PeekMessage string@, integer, integer, integer, integer  
	DECLARE integer TranslateMessage IN WIN32API  AS WinAPI_TranslateMessage string@  
	DECLARE integer DispatchMessage IN WIN32API  AS WinAPI_DispatchMessage string@
ENDIF

	IF _vfp.AutoYield   
		* В этом случае можно положиться на VFP  
		DOEVENTS  
	ELSE  
		* Сделаем за VFP его работу - обработаем системную очередь сообщений	  
		m.cWM = Space(28)  
		DO WHILE ( WinAPI_PeekMessage( @cWM, 0, 0, 0, 1 ) <> 0 )  
			WinAPI_TranslateMessage( @cWM )  
			WinAPI_DispatchMessage( @cWM )  
		ENDDO  
		* А теперь во внутренней очереди VFP есть сообщения  
		DOEVENTS  
	ENDIF  
ENDFUNC


* Покажем список SQL серверов с использованием SQLDMO
FUNCTION EnumerateSQLSERVER
	CLEAR
	oSQLApp = CREATEOBJECT('SQLDMO.Application')
	oSqlServerList = oSQLApp.ListAvailableSQLServers()
	FOR i = 1 TO oSqlServerList.COUNT
		=oSqlServerList.ITEM(i)
	ENDFOR
ENDFUNC

*
* Определение версии EXE-файла
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
* Действие над файлами
*  tnHWND - дискриптор окна
*  tcFrom - спецификация исходных файлов (можно указывать звёздочки и "?")  
*  tcTo   - спецификация результирующих файлов (папки)  
*  tnOper - определяет тип операции.  
*	     tnOper=1 - переместить файлы  
*	     tnOper=2 - копировать файлы  
*	     tnOper=3 - удалить файлы  
*	     tnOper=4 - переименовать файл  
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
 * Начинаем формировать структуру в переменной lcSHFO  
 * 
 	*_screen.HWnd
  * lcSHFO = BINTOC(thisform.HWnd, '4RS')      && Поле hwnd структуры  
  lcSHFO = BINTOC(tnHWND, '4RS')      && Поле hwnd структуры
  lcSHFO = lcSHFO + BINTOC(tnOper, '4RS')    && Поле wFunc - вид операции   
 *  
 * Обработка спецификации исходных файлов  
 *  
  tcFrom = tcFrom + CHR(0) + CHR(0)             && Дописываем нули  
  lnLenFrom = LEN(tcFrom)                       && Длина исходной строки  
  hGlobalFrom = GlobalAlloc(0x0040, lnLenFrom)  && Выделяем для неё  
                                                && блок памяти  
  SYS(2600, hGlobalFrom, lnLenFrom, tcFrom)     && и копируем туда строку  
  lcSHFO = lcSHFO + BINTOC(hGlobalFrom, '4RS')  && Поле pFrom  
 *  
 * Обработка спецификации результирующих файлов  
 *  
  IF tnOper = 3  
      lcSHFO = lcSHFO + BINTOC(0, '4RS') && Если операции удаления  
  ELSE  
      tcTo = tcTo + CHR(0) + CHR(0)      && Дописываем нули  
      lnLenTo = LEN(tcTo)                && Длина результирующей строки  
      hGlobalTo = GlobalAlloc(0x0040, lnLenTo)   && Выделяем для неё  
                                                 && блок памяти  
      SYS(2600, hGlobalTo, lnLenTo, tcTo)        && и копируем туда строку  
      lcSHFO = lcSHFO + BINTOC(hGlobalTo, '4RS') && Поле pTo структуры  
  ENDIF  
  lnFlag = 8 + 512  
  lcSHFO = lcSHFO + BINTOC(lnFlag, '2RS') && Поле fFlags структуры  
  lcSHFO = lcSHFO + REPLICATE(CHR(0), 12) && Последние 3 поля структуры  
 *  
 * Выполняем операцию  
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
 * Возвращаем память Windows  
 *  
  GlobalFree(hGlobalFrom)  
  IF tnOper != 3  
      GlobalFree(hGlobalTo)  
  ENDIF  
 *  
 * Если lnReturn = 0, то операция завершена успешно  
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
* Клас формирования ФИО
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
* Деление фамилии на составные части
*
* lcFio - Фамилия Имя Отчество
* liMode - 1-Фио 2-Имя 3-Отчество
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
* СУММА ПРОПИСЬЮ Visual FoxPro
*
* Принимает число от 0 до 999'999'999'999,99
* Возвращает текстовую строку с суммой в рублях
*
* Если число отрицательное, то берет модуль числа.
* По желании не сложно доработать, чтобы второй параметр указывал тип валюты,
* и вместо "рублей" и "копеек" подставлять соответствующие слова.
*
* Оптимизация по скорости для многократного вызова в цикле не проводилась.
********************************************************************************************
* nValuta = 0 - ничего не печатает
*         = 1 - рубли
*         = 2 - доллары
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
Asw1="сто "
case substr(Asw,1,1)="2"
Asw1="двести "
case substr(Asw,1,1)="3"
Asw1="триста "
case substr(Asw,1,1)="4"
Asw1="четыреста "
case substr(Asw,1,1)="5"
Asw1="пятьсот "
case substr(Asw,1,1)="6"
Asw1="шестьсот "
case substr(Asw,1,1)="7"
Asw1="семьсот "
case substr(Asw,1,1)="8"
Asw1="восемьсот "
case substr(Asw,1,1)="9"
Asw1="девятьсот "
endcase
cRet=cRet+Asw1
Asw1=""
if substr(Asw,2,1)="1"
do case
case substr(Asw,3,1)="0"
Asw1="десять "
case substr(Asw,3,1)="1"
Asw1="одиннадцать "
case substr(Asw,3,1)="2"
Asw1="двенадцать "
case substr(Asw,3,1)="3"
Asw1="тринадцать "
case substr(Asw,3,1)="4"
Asw1="четырнадцать "
case substr(Asw,3,1)="5"
Asw1="пятнадцать "
case substr(Asw,3,1)="6"
Asw1="шестнадцать "
case substr(Asw,3,1)="7"
Asw1="семнадцать "
case substr(Asw,3,1)="8"
Asw1="восемнадцать "
case substr(Asw,3,1)="9"
Asw1="девятнадцать "
endcase
cRet=cRet+Asw1
do case
case AswCounter=1
cRet=cRet+"миллиардов "
case AswCounter=4
cRet=cRet+"миллионов "
case AswCounter=7
cRet=cRet+"тысяч "
case AswCounter=10
	DO CASE
	CASE nValuta = 0
		cRet=cRet+" "
	CASE nValuta = 1
		cRet=cRet+"рублей "
	CASE nValuta = 2
		cRet=cRet+"USD "
	ENDCASE
endcase
else
do case
case substr(Asw,2,1)="2"
Asw1="двадцать "
case substr(Asw,2,1)="3"
Asw1="тридцать "
case substr(Asw,2,1)="4"
Asw1="сорок "
case substr(Asw,2,1)="5"
Asw1="пятьдесят "
case substr(Asw,2,1)="6"
Asw1="шестьдесят "
case substr(Asw,2,1)="7"
Asw1="семьдесят "
case substr(Asw,2,1)="8"
Asw1="восемьдесят "
case substr(Asw,2,1)="9"
Asw1="девяносто "
endcase
cRet=cRet+Asw1
Asw1=""
do case
case substr(Asw,1,3)="000"
Asw1=iif(AswCounter=10,"рублей ","")
case substr(Asw,3,1)="0" .and. substr(Asw,1,3)#"000"
Asw1=iif(AswCounter=7,"тысяч ", ;
iif(AswCounter=1,"миллиардов ",iif(AswCounter=4,"миллионов ","рублей ")))
case substr(Asw,3,1)="1"
Asw1=iif(AswCounter=7,"одна тысяча ", ;
"один "+iif(AswCounter=1,"миллиард ",iif(AswCounter=4,"миллион ","рубль ")))
case substr(Asw,3,1)="2"
Asw1=iif(AswCounter=7,"две тысячи ", ;
"два "+iif(AswCounter=1,"миллиарда ",iif(AswCounter=4,"миллиона ","рубля ")))
case substr(Asw,3,1)="3"
Asw1=iif(AswCounter=7,"три тысячи ", ;
"три "+iif(AswCounter=1,"миллиарда ",iif(AswCounter=4,"миллиона ","рубля ")))
case substr(Asw,3,1)="4"
Asw1=iif(AswCounter=7,"четыре тысячи ", ;
"четыре "+iif(AswCounter=1,"миллиарда ",iif(AswCounter=4,"миллиона ","рубля ")))
case substr(Asw,3,1)="5"
Asw1=iif(AswCounter=7,"пять тысяч ", ;
"пять "+iif(AswCounter=1,"миллиардов ",iif(AswCounter=4,"миллионов ","рублей ")))
case substr(Asw,3,1)="6"
Asw1=iif(AswCounter=7,"шесть тысяч ", ;
"шесть "+iif(AswCounter=1,"миллиардов ",iif(AswCounter=4,"миллионов ","рублей ")))
case substr(Asw,3,1)="7"
Asw1=iif(AswCounter=7,"семь тысяч ", ;
"семь "+iif(AswCounter=1,"миллиардов ",iif(AswCounter=4,"миллионов ","рублей ")))
case substr(Asw,3,1)="8"
Asw1=iif(AswCounter=7,"восемь тысяч ", ;
"восемь "+iif(AswCounter=1,"миллиардов ",iif(AswCounter=4,"миллионов ","рублей ")))
case substr(Asw,3,1)="9"
Asw1=iif(AswCounter=7,"девять тысяч ", ;
"девять "+iif(AswCounter=1,"миллиардов ",iif(AswCounter=4,"миллионов ","рублей ")))
endcase
cRet=cRet+Asw1
endif
endif
ENDFOR
AswS=substr(AswS,14,2)

if substr(AswS,1,1)="1" .or. substr(AswS,2,1)="0"
cRet=cRet+AswS+" копеек"
else
do case
case substr(AswS,2,1)>"4"
cRet=cRet+AswS+" копеек"
case substr(AswS,2,1)="1"
cRet=cRet+AswS+" копейка"
case substr(AswS,2,1)>"1" .and. substr(AswS,2,1)<"5"
cRet=cRet+AswS+" копейки"
endcase
endif
cRet=upper(substr(cRet,1,1))+substr(cRet,2)
if nSum<1
cRet="Ноль "+cRet
endif
return cRet


* Function GetProxySettings  
* Источник: forum.foxclub.ru  
* На основе кода: www.tech-archive.net  
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
  ? "Proxy и порт: "+lProxyAndPort  
    
  RETURN
ENDFUNC

*
* Function: INPUT_MSG(m.cInputPrompt, m.cDialogCaption, m.cDefaultValue, m.cInputFormat)
*
* m.cInputPrompt - текст отображающийся над полем для ввода
* m.cDialogCaption - заголовок окна
* m.cDefaultValue - значение по умолчанию
* m.cInputFormat - формат, "K","D","T","999,99" или любой другой
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
   * k - копейки  
    cSumma = TRANSFORM(M.nSumma,'9,9,,9,,,,,,9,9,,9,,,,,9,9,,9,,,,9,9,,9,,,.99')+'k'  

   * t - тысячи; m - миллионы; M - миллиарды  
    cSumma = STRTRAN(M.cSumma, ',,,,,,', 'eM')  
    cSumma = STRTRAN(M.cSumma, ',,,,,',  'em')  
    cSumma = STRTRAN(M.cSumma, ',,,,',   'et')  
    
   * e - единицы; d - десятки; c - сотни  
    cSumma = STRTRAN(M.cSumma, ',,,', 'e')  
    cSumma = STRTRAN(M.cSumma, ',,',  'd')  
    cSumma = STRTRAN(M.cSumma, ',',   'c')  
    
    cSumma = STRTRAN(M.cSumma, '0c0d0et', '')  
    cSumma = STRTRAN(M.cSumma, '0c0d0em', '')  
    cSumma = STRTRAN(M.cSumma, '0c0d0eM', '')  
    
    cSumma = STRTRAN(M.cSumma, '0c', '')  
    cSumma = STRTRAN(M.cSumma, '1c', 'сто ')  
    cSumma = STRTRAN(M.cSumma, '2c', 'двести ')  
    cSumma = STRTRAN(M.cSumma, '3c', 'триста ')  
    cSumma = STRTRAN(M.cSumma, '4c', 'четыреста ')  
    cSumma = STRTRAN(M.cSumma, '5c', 'пятьсот ')  
    cSumma = STRTRAN(M.cSumma, '6c', 'шестьсот ')  
    cSumma = STRTRAN(M.cSumma, '7c', 'семьсот ')  
    cSumma = STRTRAN(M.cSumma, '8c', 'восемьсот ')  
    cSumma = STRTRAN(M.cSumma, '9c', 'девятьсот ')  
    
    cSumma = STRTRAN(M.cSumma, '1d0e', 'десять ')  
    cSumma = STRTRAN(M.cSumma, '1d1e', 'одиннадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d2e', 'двенадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d3e', 'тринадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d4e', 'четырнадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d5e', 'пятнадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d6e', 'шестнадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d7e', 'семьнадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d8e', 'восемнадцать ')  
    cSumma = STRTRAN(M.cSumma, '1d9e', 'девятнадцать ')  
 
    cSumma = STRTRAN(M.cSumma, '0d', '')  
    cSumma = STRTRAN(M.cSumma, '2d', 'двадцать ')  
    cSumma = STRTRAN(M.cSumma, '3d', 'тридцать ')  
    cSumma = STRTRAN(M.cSumma, '4d', 'сорок ')  
    cSumma = STRTRAN(M.cSumma, '5d', 'пятьдесят ')  
    cSumma = STRTRAN(M.cSumma, '6d', 'шестьдесят ')  
    cSumma = STRTRAN(M.cSumma, '7d', 'семьдесят ')  
    cSumma = STRTRAN(M.cSumma, '8d', 'восемьдесят ')  
    cSumma = STRTRAN(M.cSumma, '9d', 'девяносто ')  
    
    cSumma = STRTRAN(M.cSumma, '0e', '')  
    cSumma = STRTRAN(M.cSumma, '5e', 'пять ')  
    cSumma = STRTRAN(M.cSumma, '6e', 'шесть ')  
    cSumma = STRTRAN(M.cSumma, '7e', 'семь ')  
    cSumma = STRTRAN(M.cSumma, '8e', 'восемь ')  
    cSumma = STRTRAN(M.cSumma, '9e', 'девять ')  

***************************************************************
    cSumma = STRTRAN(M.cSumma, '1e.', 'один ')  
    cSumma = STRTRAN(M.cSumma, '2e.', 'два ')  
    cSumma = STRTRAN(M.cSumma, '3e.', 'три ')  
    cSumma = STRTRAN(M.cSumma, '4e.', 'четыре ')  
    cSumma = STRTRAN(M.cSumma, '1et', 'одна тысяча ')  
    cSumma = STRTRAN(M.cSumma, '2et', 'две тысячи ')  
    cSumma = STRTRAN(M.cSumma, '3et', 'три тысячи ')  
    cSumma = STRTRAN(M.cSumma, '4et', 'четыре тысячи ')  
    cSumma = STRTRAN(M.cSumma, '1em', 'один миллион ')  
    cSumma = STRTRAN(M.cSumma, '2em', 'два миллиона ')  
    cSumma = STRTRAN(M.cSumma, '3em', 'три миллиона ')  
    cSumma = STRTRAN(M.cSumma, '4em', 'четыре миллиона ')  
    cSumma = STRTRAN(M.cSumma, '1eM', 'один миллиард ')  
    cSumma = STRTRAN(M.cSumma, '2eM', 'два миллиарда ')  
    cSumma = STRTRAN(M.cSumma, '3eM', 'три миллиарда ')  
    cSumma = STRTRAN(M.cSumma, '4eM', 'четыре миллиарда ')  
    
    cSumma = STRTRAN(M.cSumma, '11k', 'одиннадцать ')  
    cSumma = STRTRAN(M.cSumma, '12k', 'двенадцать ')  
    cSumma = STRTRAN(M.cSumma, '13k', 'тринадцать ')  
    cSumma = STRTRAN(M.cSumma, '14k', 'четырнадцать ')  
    cSumma = STRTRAN(M.cSumma, '1k', 'один ')  
    cSumma = STRTRAN(M.cSumma, '2k', 'два ')  
    cSumma = STRTRAN(M.cSumma, '3k', 'три ')  
    cSumma = STRTRAN(M.cSumma, '4k', 'четыре ')  
    
    cSumma = STRTRAN(M.cSumma, '.', '')  
    cSumma = STRTRAN(M.cSumma, 't', 'тысяч ')  
    cSumma = STRTRAN(M.cSumma, 'm', 'миллионов ')  
    cSumma = STRTRAN(M.cSumma, 'M', 'миллиардов ')  
    cSumma = STRTRAN(M.cSumma, 'k', '')  
    cSumma = STRTRAN(M.cSumma, ' 00', '')  
***************************************************************
    
*!*	    cSumma = STRTRAN(M.cSumma, '1e.', 'один рубль ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2e.', 'два рубля ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3e.', 'три рубля ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4e.', 'четыре рубля ')  
*!*	    cSumma = STRTRAN(M.cSumma, '1et', 'одна тысяча ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2et', 'две тысячи ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3et', 'три тысячи ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4et', 'четыре тысячи ')  
*!*	    cSumma = STRTRAN(M.cSumma, '1em', 'один миллион ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2em', 'два миллиона ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3em', 'три миллиона ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4em', 'четыре миллиона ')  
*!*	    cSumma = STRTRAN(M.cSumma, '1eM', 'один миллиард ')  
*!*	    cSumma = STRTRAN(M.cSumma, '2eM', 'два миллиарда ')  
*!*	    cSumma = STRTRAN(M.cSumma, '3eM', 'три миллиарда ')  
*!*	    cSumma = STRTRAN(M.cSumma, '4eM', 'четыре миллиарда ')  
*!*	    
*!*	    cSumma = STRTRAN(M.cSumma, '11k', '11 копеек')  
*!*	    cSumma = STRTRAN(M.cSumma, '12k', '12 копеек')  
*!*	    cSumma = STRTRAN(M.cSumma, '13k', '13 копеек')  
*!*	    cSumma = STRTRAN(M.cSumma, '14k', '14 копеек')  
*!*	    cSumma = STRTRAN(M.cSumma, '1k', '1 копейка')  
*!*	    cSumma = STRTRAN(M.cSumma, '2k', '2 копейки')  
*!*	    cSumma = STRTRAN(M.cSumma, '3k', '3 копейки')  
*!*	    cSumma = STRTRAN(M.cSumma, '4k', '4 копейки')  
*!*	    
*!*	    cSumma = STRTRAN(M.cSumma, '.', 'рублей ')  
*!*	    cSumma = STRTRAN(M.cSumma, 't', 'тысяч ')  
*!*	    cSumma = STRTRAN(M.cSumma, 'm', 'миллионов ')  
*!*	    cSumma = STRTRAN(M.cSumma, 'M', 'миллиардов ')  
*!*	    cSumma = STRTRAN(M.cSumma, 'k', ' копеек')  
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
  m.e900='сто, двести, триста, четыреста, пятьсот, шестьсот, семьсот, восемьсот, девятьсот,'  
  m.e90='десять, двадцать, тридцать, сорок, пятьдесят, шестьдесят, семьдесят, восемьдесят, девяносто,'  
  m.e19='одиннадцать, двенадцать, тринадцать, четырнадцать, пятнадцать, шестнадцать, семнадцать, восемнадцать, девятнадцать,'  
  m.e10m='один, два, три, четыре, пять, шесть, семь, восемь, девять,'  
  m.e10z='одна, две, три, четыре, пять, шесть, семь, восемь, девять,'  
  m.m='тысяч миллион миллиард триллион квадриллион квинтиллион секстиллион септиллион октиллион нониллион дециллион'  
  m.v=''  
  FOR lni=m.r3 TO 1 STEP -1  
  	m.c123=SUBSTR(m.cSumma,m.nPoint-3*lni,3)  
  	m.n12=IIF(BETWEEN(RIGHT(m.c123,2),'11','19'),VAL(RIGHT(m.c123,1)),0)  
  	m.n2=IIF(m.n12=0,VAL(SUBSTR(m.c123,2,1)),0)  
  	m.n1=IIF(m.n12=0 and m.n2#1,VAL(RIGHT(m.c123,1)),0)  
  	m.n3=VAL(LEFT(m.c123,1))  
  	m.v123=GETWORDNUM(m.e900,n3)+GETWORDNUM(m.e19,m.n12)+GETWORDNUM(m.e90,m.n2)+GETWORDNUM(IIF(lni=2,m.e10z,m.e10m),m.n1)  
  	m.v=m.v+m.v123+IIF(lni=1,'',IIF(VAL(m.c123)>0,GETWORDNUM(m.m,lni-1)+IIF(m.n1=1,IIF(lni=2,'a',''),IIF(BETWEEN(m.n1,2,4),IIF(lni=2,'и','а'),IIF(lni=2,'','ов')))+',',''))  
  ENDFOR  
  m.v=IIF(EMPTY(m.v),'ноль ',m.v)  
 *RETURN CHRTRAN(PROPER(m.v),',',' ')+SUBSTR(m.cSumma,m.nPoint+1,2)  
  m.cDrob=SUBSTR(m.cSumma,m.nPoint+1,2)  
  m.cBig='рубл'+ICASE(m.n12>0 OR m.n1=0 OR BETWEEN(m.n1,5,9),'ей',m.n12=0 AND BETWEEN(m.n1,2,4),'я','ь')+' '  
  m.cSmall=' копе'+ICASE(LEFT(m.cDrob,1)='1' OR RIGHT(m.cDrob,1)='0' OR BETWEEN(RIGHT(m.cDrob,1),'5','9'),'ек',BETWEEN(RIGHT(m.cDrob,1),'2','4'),'йки','йка')  
  RETURN CHRTRAN(PROPER(m.v),',',' ') + m.cBig + m.cDrob + m.cSmall
ENDFUNC

FUNCTION ErrorHandler(nError,cMethod,nLine)
	DO WHILE TXNLEVEL( )>0
		ROLLBACK
	ENDDO

	DO CASE
	CASE nError = 1429
		MESSAGEBOX("#1429.SQL Server не существует, или доступ запрещен..."+CHR(13)+"Обратитесь к системному администратору для выяснения причины"+CHR(13)+MESSAGE(), 48, "Ошибка подключения к SQL-серверу")
		ON ERROR
		QUIT

	CASE nError = 1435
		MESSAGEBOX("#1435.Ошибка в структуре данных"+CHR(13)+MESSAGE()+CHR(13)+"Сообщите об данной ошибке администратору",48,_vfp.Caption)
		ON ERROR
		QUIT
	
	CASE nError = 2005
		MESSAGEBOX("#2005.Ошибка в структуре данных"+CHR(13)+MESSAGE()+CHR(13)+"Сообщите об данной ошибке администратору",48,_vfp.Caption)
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
Дата/время возникновения ошибки: <<TTOC(DATETIME())>><br>
Программый стэк:<br>
<<lcPrg>><br>
Метод вызвавший ошибку: <<cMethod>><br>
Строка: <<TRANSFORM(nLine)>><br>
Номер ошибки: <<TRANSFORM(nError)>><br>
Сообщение: <<MESSAGE()>><br>
Дополнительное сообщение: <<MESSAGE(1)>><br>
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
		SetSubject("Ошибка в программе")
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
			SYS(0)+". Ошибка в программе",;
			lcBody,;
			"",;
			lcFile)
	ENDIF

	IF MESSAGEBOX("Возникла ошибка в программе"+lcCrLf+"Информация направлена разработчику"+lcCrLf+"Подробности в файле"+lcCrLf+lcFile+lcCrLf+lcCrLf+"Продолжить работу программы?",17,_screen.Caption)#1
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
* Класс для подключения к SQL серверу
* Параметр: хендлер подключения полученный в результате sqlstringconnect()
*
DEFINE CLASS clSQLconnect as Custom
	SQL_ConString = []
	SQL_Comment = [Пустой комментарий]
	SQL_ConnHandler = -1
	SQL_ErrorMessage = []
	SQL_ErrorNumber = 0
	SQL_FileLog = "D:\1\sql.txt"
	CRLF = CHR(13)+CHR(10)
	
	PROCEDURE Init
		LPARAMETERS liH as Integer			&& Хендлер
		
		This.SQL_ConnHandler = liH
		IF This.SQL_ConnHandler <= 0
			MESSAGEBOX("Подключение к SQL-серверу не выполнено: Handler="+TRANSFORM(This.SQL_ConnHandler)+CHR(13)+"Дальнейшая работа приложения не возможна",48,THIS.Name)
			RETURN .F.
		ENDIF
		*STRTOFILE("Подключение"+THIS.CRLF, THIS.SQL_FileLog)
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
		*This.SQL_Log.Puts("Выполнение SQL-запроса>>")
		*This.SQL_Log.Puts(lcSQL1)
		LOCAL liResult as Integer
		IF TYPE("lcAlias")="C" AND USED(lcAlias) AND CURSORGETPROP("Buffering", lcAlias) > 3
			TABLEREVERT(.T., lcAlias)
		ENDIF
		liResult = SQLEXEC(This.SQL_ConnHandler, lcSQL1, lcAlias)
		
		*STRTOFILE("Выполнение:"+THIS.CRLF, THIS.SQL_FileLog,1)
		*STRTOFILE(lcSQL1+THIS.CRLF, THIS.SQL_FileLog,1)
		*STRTOFILE("Результат: "+TRANSFORM(liResult)+THIS.CRLF, THIS.SQL_FileLog,1)
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

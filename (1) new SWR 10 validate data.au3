#include "S:\PERMITTING\ENG\2008 Eng Procedures\Survey Technical Review\Directional survey deficiency letters\letter program\mainframe functions.au3"
#include <Array.au3>
#include <word.au3>
#Include <Date.au3>

$scripttitle="SWR 10 validate data script 234623624746"
while winexists($scripttitle)
	winkill($scripttitle)
wend
autoitwinsettitle($scripttitle)
opt("winwaitdelay", 1)
opt("wintitlematchmode", 2)
opt("sendkeydelay", 10)
opt("trayicondebug", 1)


$hostdirectory=@scriptdir & "\"
;$hostdirectory="C:\Documents and Settings\RosenquS\Desktop\bak\Rule 10\SWR 10\"

$temp=stringsplit("%userprofile%", "\")
$username=$temp[$temp[0]]
$e = ObjGet("", "Excel.Application")

$workbookname=""
$temp=processlist("excel.exe")
if $temp[0][0]>=1 then
	for $i=2 to $temp[0][0]
		processclose($temp[$i][1])
	next
endif
$found=0
dim $workbookobj
do
	$e = ObjGet("", "Excel.Application")
	while not isobj($e)
		$answer=msgbox(6+262144, "", "Open a SWR 10 data spreadsheet.  Click CANCEL to quit.  Click TRY AGAIN to TERMINATE all Excel instances.  Click CONTINUE to keep going.")
		if $answer=2 then exit
		if $answer=10 then runwait("taskkill.exe /f /im excel.exe")
		;if $answer=2 then runwait("taskkill.exe /f /im excel.exe")
		$e = ObjGet("", "Excel.Application")
	wend		
	if isobj($e) then
		$e.visible=1	
		for $tempworkbookobj in $e.workbooks
			if stringinstr($tempworkbookobj.name, "SWR 10 data (") then
				$found=1
				$workbookname=$tempworkbookobj.name
				$workbookobj=$tempworkbookobj
			endif
		next
		if $found=0 then
			$answer=msgbox(6+262144, "", "Open a SWR 10 data spreadsheet.  Click CANCEL to quit.  Click TRY AGAIN to TERMINATE all Excel instances.  Click CONTINUE to keep going.")
			if $answer=2 then exit
			if $answer=10 then runwait("taskkill.exe /f /im excel.exe")
		endif
	endif
until $found=1

$sheetobj=$workbookobj.worksheets("Sheet1")
$workbookobj.activate
$sheetobj.activate

;$e.workbooks($workbookname).activate
;$sheetobj.activate

$currentrow=$e.activecell.row




$sheetobj.rows($currentrow-1).interior.colorindex=-4142
$sheetobj.rows($currentrow).interior.colorindex=4
$workbookname=$e.activeworkbook.name

$firstrow=$e.selection.row
$lastrow=1
for $rowobject in $e.selection.rows
	$lastrow=$rowobject.row
next

for $currentrow=$firstrow to $lastrow

if $workbookobj.path<>"S:\PERMITTING\ENG\SWR 10\SWR 10" then
	msgbox(262144, "", "ERROR: the workbook file location is """ & $workbookobj.path & """ instead of ""S:\PERMITTING\ENG\SWR 10\SWR 10""")
	exit
endif
$sheetobj.cells($currentrow, getcol($sheetobj, "API#")).activate




dim $fieldnames[100]
dim $fieldnumbers[100]
dim $numfields=0, $api="", $dp, $letterdate, $lease, $well, $letterdate_dayvalue, $approvaldate_dayvalue, $approvaldate_longdate, $letterdate_longdate, $fieldname, $h2s, $receivedate, $dpissueddate, $spuddate, $surfcasingdate
;$currentrow=890
;$currentrow=$e.activecell.row
tooltip("current row #" & $currentrow)

;findlastrow($currentrow)

$attn=$sheetobj.cells($currentrow, getcol($sheetobj, "ATTN")).text

$api=$sheetobj.cells($currentrow, getcol($sheetobj, "API#")).value
while stringlen($api)<>9
	$sheetobj.cells($currentrow, getcol($sheetobj, "API#")).activate
	$answer=msgbox(1+262144, "", "enter API number" & @CRLF & "excel row #" & $currentrow)
	if $answer=2 then exit
	$api=$sheetobj.cells($currentrow, getcol($sheetobj, "API#")).value
wend

$receivedate=$sheetobj.cells($currentrow, getcol($sheetobj, "receive date")).text
while $receivedate=""
	$sheetobj.cells($currentrow, getcol($sheetobj, "receive date")).activate
	$answer=msgbox(1+262144, "", "enter receive date" & @CRLF & "excel row #" & $currentrow)
	if $answer=2 then exit
	$receivedate=$sheetobj.cells($currentrow, getcol($sheetobj, "receive date")).text
wend
$blanket=$sheetobj.cells($currentrow, getcol($sheetobj, "blanket (yes or no)")).value
while not (stringinstr($blanket, "yes") or stringinstr($blanket, "no"))
	$sheetobj.cells($currentrow, getcol($sheetobj, "blanket (yes or no)")).activate
	$answer=msgbox(1+262144, "", "blanket (yes or no)" & @CRLF & "excel row #" & $currentrow)
	if $answer=2 then exit
	$blanket=$sheetobj.cells($currentrow, getcol($sheetobj, "blanket (yes or no)")).text

wend

if $blanket="yes" then $sheetobj.cells($currentrow, getcol($sheetobj, "Docket #")).value=""

if $blanket="no" then
	$TOCprod=$sheetobj.cells($currentrow, getcol($sheetobj, "production string TOC")).value
	while $TOCprod=""
		$sheetobj.cells($currentrow, getcol($sheetobj, "production string TOC")).activate
		$answer=msgbox(1+262144, "", "must enter TOC if the application is non-blankets" & @CRLF & "excel row #" & $currentrow)
		if $answer=2 then exit
		$TOCprod=$sheetobj.cells($currentrow, getcol($sheetobj, "production string TOC")).value
	wend
endif

;msgbox(1+262144, "", $api & " ROYALTY AND WORKING INTERESTS")
wait()

mainframetype(2, 2, "WBTM{enter}")
$screen=mainframegetscreen("WELL BORE TECHNICAL DATA MENU")
mainframetype(8, 23, stringmid($api, 1, 3))
mainframetype(8, 27, stringmid($api, 5, 5))
mainframetype(11, 8, "s")
mainframetype(14, 7, "s{enter}")


$permithistoryscreen=mainframegetscreen("PERMIT NUMBERS AND WELLS WITHIN WELL BORE")
$currentrowow=5
$foundpermit=0
do
	$currentrowow+=1
	$permit=mainframecopy($permithistoryscreen, $currentrowow, 7, 6)
	if stringlen($permit)=6 and stringisdigit($permit) then $foundpermit=1
until stringinstr($permit, "___") or stringlen($permit)<=1
if $foundpermit=1 then
	mainframetype($currentrowow-1, 4, "s")
	mainframetype(18, 4, "s{enter}")

	$permitscreen=mainframegetscreen("DRILLING PERMIT MASTER DATA INQUIRY")
	$opname=mainframecopy($permitscreen, 5, 17, 34)
	$lease=mainframecopy($permitscreen, 9, 17, 34)
	$district=mainframecopy($permitscreen, 10, 17, 2)
	$county=""
	$dpissueddate=mainframecopy($permitscreen, 15, 17, 10)
	$dpissueddate=stringreplace($dpissueddate, " ", "-")
	$dp=mainframecopy($permitscreen, 3, 17, 6)
	$spuddate=mainframecopy($permitscreen, 19, 42, 10)
	$surfcasingdate=mainframecopy($permitscreen, 19, 17, 10)
	if stringmid($api, 1, 1)="6" or stringmid($api, 1, 1)="7" then
		$county=mainframecopy($permitscreen, 12, 17, 34)
	else
		$county=mainframecopy($permitscreen, 11, 17, 34)
	endif
	$well=mainframecopy($permitscreen, 9, 69, 10)
	$opnumber=mainframecopy($permitscreen, 5, 69, 6)
	$api=mainframecopy($permitscreen, 12, 69, 3) & "-" & mainframecopy($permitscreen, 12, 73, 5)
	if stringinstr($permitscreen, "PERMIT TYPE => DRILL") then
		$sheetobj.cells($currentrow, getcol($sheetobj, "new drill (yes/no)")).value="yes"
	else
		$sheetobj.cells($currentrow, getcol($sheetobj, "new drill (yes/no)")).value="no"
	endif
else
	mainframetype(16, 4, "s@B@Bs{enter}")

	$W2G1screen=mainframegetscreen("OIL AND GAS W-2/G-1 RECORD")
	$opname=mainframecopy($W2G1screen, 5, 9, 34)
	$lease=mainframecopy($W2G1screen, 4, 48, 31)
	$api=mainframecopy($W2G1screen, 2, 9, 3) & "-" & mainframecopy($W2G1screen, 2, 15, 5)
	$district=mainframecopy($W2G1screen, 3, 9, 2)
	$county=mainframecopy($W2G1screen, 3, 66, 14)
	$well=mainframecopy($W2G1screen, 3, 36, 9)

	wait2()

	mainframetype(10, 10, "ornq " & $opname & "{enter}")


	$screen=mainframegetscreen("ORGANIZATION NAME INQUIRY")
	$opnumber=mainframecopy($screen, 4, 42, 6)
endif
;send("{enter}")



$wait=msgbox(1+262144, "", $opname & @CRLF & @CRLF & "click OK to change operator number", 2)
if $wait=1 then

	do
		$opnumber=inputbox("", "confirm operator number" & @CRLF & @CRLF & $opname, $opnumber)
		if $opnumber="" then
			$sheetobj.rows($currentrow).interior.colorindex=-4142
			exit
		endif
	until stringisdigit($opnumber) and stringlen($opnumber)=6
endif
dim $line1[100], $line2[100], $line3[100]

$ormqscreenpath=@scriptdir & "\ormq screens\" & $opnumber & ".txt"
$filemodifieddate=filegettime($ormqscreenpath)
$p5screen=""
if (not fileexists($ormqscreenpath)) or (_datetodayvalue($filemodifieddate[0], $filemodifieddate[1], $filemodifieddate[2])+14 < _datetodayvalue(@year, @mon, @mday)) then
	wait2()
	mainframetype(2, 2, "ormq " & $opnumber & "{enter}")


	$p5screen=mainframegetscreen("OPERATOR NUMBER")
	$output=fileopen($ormqscreenpath, 2)
	filewrite($output, $p5screen)
	fileclose($output)
else
	$input=fileopen($ormqscreenpath, 0)
	$p5screen=fileread($input)
	fileclose($input)
endif

$opnumber=mainframecopy($p5screen, 3, 19, 6)
$opname=mainframecopy($p5screen, 4, 12, 34)
$line1[0]=mainframecopy($p5screen, 6, 3, 38)
$line2[0]=mainframecopy($p5screen, 7, 3, 38)
$line3[0]=mainframecopy($p5screen, 8, 3, 38)



$index=0
if not stringinstr($attn, "@") then
	$screen=""
    wait2()
	mainframetype(2, 2, "ORAR " & $opnumber & "{enter}")

	$i=1

	do
		$screen=mainframegetscreen("ADDRESS INQUIRY")
		mainframetype(1, 1, "{f5}")
		sleep(100)
		$screen=mainframegetscreen("ADDRESS INQUIRY")
		;mainframewaitchange($screen)
		for $r=4 to 16 step 4
			for $c=2 to 44 step 42
				$line1[$i]=mainframecopy($screen, $r, $c, 29)
				$line2[$i]=mainframecopy($screen, $r+1, $c, 29)
				$line3[$i]=mainframecopy($screen, $r+2, $c, 29)
				$i+=1
			next
		next
	until stringinstr($screen, "NO MORE ADDRESSES")
	$numaddresses=$i

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	$prompt=""
	for $i=0 to $numaddresses-1
		$prompt &= $i & ".	"  & $line1[$i] & "		" & $line2[$i] & "		" & $line3[$i] & @CRLF
		if mod($i+1, 5)=0 then $prompt &=@CRLF
	next

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	$index=0

	$attn=uppercase($attn)
	$index=inputbox("Select Address", $prompt, "0", "", 700, 700)
	if $index="" then $index=0
endif

$opaddress1=$line1[$index]
$opaddress2=$line2[$index]
$opaddress3=$line3[$index]

;$sheetobj.cells($currentrow, getcol($sheetobj, "ATTN")).value=$attn

if (stringinstr($opaddress1, "ATTN") or stringinstr($opaddress1, "C/O")) and $attn<>"" and not stringinstr($attn, "@") then $opaddress1="ATTN " & $attn

if $opaddress3="" then
	$opaddress3=$opaddress2
	$opaddress2=$opaddress1
	$opaddress1="ATTN: REGULATORY DEPARTMENT"
	if $attn<>"" and not stringinstr($attn, "@") then $opaddress1="ATTN " & $attn
endif
if (stringinstr($opaddress1, "ATTN") or stringinstr($opaddress1, "C/O")) and stringinstr($attn, "@") then $opaddress1="ATTN: REGULATORY DEPARTMENT"


$sheetobj.cells($currentrow, getcol($sheetobj, "OPERATOR ADDRESS 1")).value=$opaddress1
$sheetobj.cells($currentrow, getcol($sheetobj, "OPERATOR ADDRESS 2")).value=$opaddress2
$sheetobj.cells($currentrow, getcol($sheetobj, "OPERATOR ADDRESS 3")).value=$opaddress3
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

$fieldnumbercol=getcol($sheetobj, "field 1")
$h2scol=getcol($sheetobj, "h2s field 1")
$fieldnamecol=getcol($sheetobj, "field name 1")
$fieldperfscol=getcol($sheetobj, "perfs field 1")
$fieldnamedisplay=""
$h2scondition="NO"
$allfieldslettertext=""



if $sheetobj.cells($currentrow, $fieldnumbercol).text="rsp" then
	$sheetobj.cells($currentrow, $fieldnumbercol).value=85280300
	$sheetobj.cells($currentrow, $fieldnumbercol+1).value=85280900
	$sheetobj.cells($currentrow, $fieldnumbercol+2).value=85448150
endif
if $sheetobj.cells($currentrow, $fieldnumbercol).text="oxy" then
	$sheetobj.cells($currentrow, $fieldnumbercol).value=85280300
	$sheetobj.cells($currentrow, $fieldnumbercol+1).value=56378750
	$sheetobj.cells($currentrow, $fieldnumbercol+2).value=55256030
endif

if $sheetobj.cells($currentrow, $fieldnumbercol).text="c" then
	$sheetobj.cells($currentrow, $fieldnumbercol).value=85280300
	$sheetobj.cells($currentrow, $fieldnumbercol+1).value=71021430
	$sheetobj.cells($currentrow, $fieldnumbercol+2).value=69765200
	$sheetobj.cells($currentrow, $fieldnumbercol+3).value=16559500
endif

while $sheetobj.cells($currentrow, $fieldnumbercol).value <> ""
	$fieldnumber=string(stringreplace($sheetobj.cells($currentrow, $fieldnumbercol).text, " ", ""))

	if not stringisdigit($fieldnumber) then $fieldnumber=shortcutfieldnames($fieldnumber)

	if $district="7C" and $e.cells($currentrow, getcol($sheetobj, "field 1")).value=85280300 then
		$e.cells($currentrow, getcol($sheetobj, "field 1")).value=85279200
	endif

	if $district="08" and $e.cells($currentrow, getcol($sheetobj, "field 1")).value=85279200 then
		$e.cells($currentrow, getcol($sheetobj, "field 1")).value=85280300
	endif
	$sheetobj.cells($currentrow, $fieldnumbercol).value=$fieldnumber

	getfieldname($fieldnumber, $h2s, $fieldname)
	if $h2s="PRESENT" then $h2scondition="YES"
	$sheetobj.cells($currentrow, $fieldnamecol).value=$fieldname
	$sheetobj.cells($currentrow, $h2scol).value=$h2s
	$perfs=$sheetobj.cells($currentrow, $fieldperfscol).value
	while stringlen($perfs)<2 and $blanket="no"
		$sheetobj.cells($currentrow, $fieldperfscol).activate
		msgbox(1+262144, "", "enter perfs" & @CRLF & @CRLF & $fieldname)
		$perfs=$sheetobj.cells($currentrow, $fieldperfscol).value
	wend
	$fieldnumber=$sheetobj.cells($currentrow, $fieldnumbercol).value

	$fieldnamedisplay &= $fieldname & "{" & $fieldnumber & "{" & $perfs & @LF
	$allfieldslettertext &= $fieldname & "; "
	$h2scol+=1
	$fieldnumbercol+=1
	$fieldnamecol+=1
	$fieldperfscol+=1
wend
$allfieldslettertext=stringtrimright($allfieldslettertext, 2)

if $e.cells($currentrow, getcol($sheetobj, "field 1")).value=85280300 or $e.cells($currentrow, getcol($sheetobj, "field 1")).value=85279200 then
	$e.cells($currentrow, getcol($sheetobj, "Allow.")).value="515 BOPD"
endif



for $pos=stringlen($allfieldslettertext) to 1 step -1
	if stringmid($allfieldslettertext, $pos, 2)="; " then
		$allfieldslettertext=stringmid($allfieldslettertext, 1, $pos) & " and " & stringmid($allfieldslettertext, $pos+2, 99)
		exitloop
	endif
next

$fieldnamedisplay=stringtrimright($fieldnamedisplay, 1)
$sheetobj.cells($currentrow, getcol($sheetobj, "Fields")).value=$fieldnamedisplay
$sheetobj.cells($currentrow, getcol($sheetobj, "Field Assgn.")).value=$sheetobj.cells($currentrow, getcol($sheetobj, "field name 1")).value






;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;populate spreadsheet

;$template=$sheetobj.cells($currentrow, getcol($sheetobj, "type of letter")).text
;$allfields=""
;$allfields_worddisplay=""
;$unassignedfields=""
;$assignedfield=""


;if $template="Oil or Gas" then
;	$sheetobj.cells($currentrow, getcol($sheetobj, "Field Assgn.")).value=$fieldnames[0] & " for gas wells" & @LF & $fieldnames[1] & " for oil wells"
;elseif $template="Oil Conditional" then
;	$sheetobj.cells($currentrow, getcol($sheetobj, "Field Assgn.")).value=$fieldnames[0] & " if _______" & @LF & $fieldnames[1] & " if __________"
;elseif $template="Gas Conditional" then
;	$sheetobj.cells($currentrow, getcol($sheetobj, "Field Assgn.")).value=$fieldnames[0] & " if _______" & @LF & $fieldnames[1] & " if __________"
;else
;	$sheetobj.cells($currentrow, getcol($sheetobj, "Field Assgn.")).value=$fieldnames[0]
;endif
;
;$allfields=""
;for $i=0 to $numfields-1
;	$allfields &= $fieldnames[$i] & " (#" & $fieldnumbers[$i] & ")" & @LF
;next
;$allfields=stringtrimright($allfields, 1)
;$sheetobj.cells($currentrow, getcol($sheetobj, "Fields")).value=$allfields
;
;$allfields=""
;for $i=0 to $numfields-1
;	$allfields &= $fieldnames[$i] & " (#" & $fieldnumbers[$i] & ")" & @CR
;next
;$allfields=stringtrimright($allfields, 1)
;$sheetobj.cells($currentrow, getcol($sheetobj, "all fields letter header")).value=$allfields
;




;$unassignedfields=""
;for $i=1 to $numfields-1
;	$unassignedfields &= $fieldnames[$i] & "; "
;	if $i=$numfields-2 then	$unassignedfields &= "and "
;next
;$unassignedfields=stringtrimright($unassignedfields, 2)
;
;$sheetobj.cells($currentrow, getcol($sheetobj, "primary field")).value=$fieldnames[0]
;$sheetobj.cells($currentrow, getcol($sheetobj, "secondary field")).value=$fieldnames[1]
;$sheetobj.cells($currentrow, getcol($sheetobj, "unassigned fields")).value=$unassignedfields
;;msgbox(1+262144, "", $allfields)

$sheetobj.cells($currentrow, getcol($sheetobj, "disposition date")).value=@year & "/" & @mon & "/" & @mday
$sheetobj.cells($currentrow, getcol($sheetobj, "script run date")).value=@year & "/" & @mon & "/" & @mday
$sheetobj.cells($currentrow, getcol($sheetobj, "District")).value=$district
$sheetobj.cells($currentrow, getcol($sheetobj, "County")).value=$county
$sheetobj.cells($currentrow, getcol($sheetobj, "Operator")).value=$opname
$sheetobj.cells($currentrow, getcol($sheetobj, "Operator no.")).value=$opnumber
$sheetobj.cells($currentrow, getcol($sheetobj, "Lease")).value=$lease
$sheetobj.cells($currentrow, getcol($sheetobj, "Well")).value=$well
$sheetobj.cells($currentrow, getcol($sheetobj, "well name combo")).value=$lease & " — WELL NO. " & $well & ", API NO. " & $api
$sheetobj.cells($currentrow, getcol($sheetobj, "h2s restriction")).value=$h2scondition
$sheetobj.cells($currentrow, getcol($sheetobj, "drilling permit issued date")).value=$dpissueddate
$sheetobj.cells($currentrow, getcol($sheetobj, "drilling permit no.")).value=$dp
$sheetobj.cells($currentrow, getcol($sheetobj, "all fields letter body")).value=$allfieldslettertext
if $spuddate<>"" or $surfcasingdate<>"" then
	$sheetobj.cells($currentrow, getcol($sheetobj, "current permit has spud date (yes/no)")).value="yes"
else
	$sheetobj.cells($currentrow, getcol($sheetobj, "current permit has spud date (yes/no)")).value="no"
endif

$leavepermitopen=$sheetobj.cells($currentrow, getcol($sheetobj, "requested permit held open (yes/no)")).value
$newdrill=$sheetobj.cells($currentrow, getcol($sheetobj, "new drill (yes/no)")).value
$alreadyonschedule=$sheetobj.cells($currentrow, getcol($sheetobj, "all zones already on schedule (yes/no)")).value
if $blanket="yes" then $sheetobj.cells($currentrow, getcol($sheetobj, "requested permit held open (yes/no)")).value="no"

if $sheetobj.cells($currentrow, getcol($sheetobj, "new drill (yes/no)")).value<>"yes" and $blanket="no" then
	while $sheetobj.cells($currentrow, getcol($sheetobj, "all zones already on schedule (yes/no)")).value=""
		msgbox(1+262144, "", "Error: answer ""all zones already on schedule (yes/no)""")
	wend
endif


$sheetobj.cells($currentrow, getcol($sheetobj, "Expiration Date")).formula=$sheetobj.cells($currentrow, getcol($sheetobj, "disposition date")).formula+731
$sheetobj.cells($currentrow, getcol($sheetobj, "expiration condition")).value="This exception to SWR 10 will expire if not used within two (2) years from the date of this permit."

if $blanket="no" and $sheetobj.cells($currentrow, getcol($sheetobj, "new drill (yes/no)")).value="no" and $sheetobj.cells($currentrow, getcol($sheetobj, "all zones already on schedule (yes/no)")).value="no" then
	$sheetobj.cells($currentrow, getcol($sheetobj, "expiration condition")).value="This exception to SWR 10 will expire if not used within two (2) years from the original date of issuance for drilling permit no. " & $dp & "."
	$sheetobj.cells($currentrow, getcol($sheetobj, "Expiration Date")).formula=$sheetobj.cells($currentrow, getcol($sheetobj, "drilling permit issued date")).formula+731
endif

if $sheetobj.cells($currentrow, getcol($sheetobj, "requested permit held open (yes/no)")).value="yes" then
	$sheetobj.cells($currentrow, getcol($sheetobj, "expiration condition")).value="This exception to SWR 10 will expire if not used within two (2) years from the original date of issuance for drilling permit no. " & $dp & "."
	$sheetobj.cells($currentrow, getcol($sheetobj, "Expiration Date")).formula=$sheetobj.cells($currentrow, getcol($sheetobj, "drilling permit issued date")).formula+731
endif





$permitconditions=""
if $h2scondition="YES" then $permitconditions &= "The commingled well will be subject to Statewide Rule 36 (operation in hydrogen sulfide areas) because at least one of the commingled fields requires a Certificate of Compliance for Statewide Rule 36.  The well must be operated in accordance with Statewide Rule 36." & @CR & @CR
if $blanket="YES" then $permitconditions &= "The completion report for the commingled well must indicate which perforations belong to which field.  The Commission may also require a wellbore diagram to be filed with the completion report for the commingled well.  If filed, the wellbore diagram must indicate which perforations belong to which field." & @CR & @CR
if $blanket="NO" then $permitconditions &= "The completion of the commingled well must be a reasonable match with the wellbore diagram filed with the application.  Variances in completion depths are acceptable provided that these completion depths remain within the designated correlative intervals for the commingled fields.  A copy of this wellbore diagram must be filed with the completion report for the commingled well." & @CR & @CR

;$sheetobj.rows($currentrow).wraptext=0

$custompermitconditions=$sheetobj.cells($currentrow, getcol($sheetobj, "custom permit conditions")).value
if $custompermitconditions<>"" then $permitconditions &= $custompermitconditions & @CR & @CR

if stringinstr($attn, "@") then $permitconditions &= "Note: The distribution of this document will be by E-MAIL ONLY.  E-mail sent to " & $attn & "." & @CR & @CR


if $permitconditions<>"" then $sheetobj.cells($currentrow, getcol($sheetobj, "permit conditions")).value="Permit conditions:" & @CR & @CR & $permitconditions


$sheetobj.range("A" & $currentrow & ":A" & $currentrow).activate
winactivate("Excel")


$sheetobj.columns(getcol($sheetobj, "ATTN")).hyperlinks.delete



$sheetobj.cells($currentrow, getcol($sheetobj, "validate run date")).value=@year & "/" & @mon & "/" & @mday

$sheetobj.cells($currentrow, getcol($sheetobj, "disposition")).value="pending"


$sheetobj.rows($currentrow).interior.colorindex=-4142

;if $sheetobj.cells($currentrow, getcol($sheetobj, "unique ID")).value="" then $sheetobj.cells($currentrow, getcol($sheetobj, "unique ID")).value="ID" & $currentrow+random()

$workbookobj.save




next





func getfieldname($fieldnumber, byref $h2s, byref $fieldname)

	$fieldname=""
	$h2s=""
	$screen=""
	$flimscreenpath=@scriptdir & "\flim screens\" & $fieldnumber & ".txt"
	$filemodifieddate=filegettime($flimscreenpath)
	if stringlen($fieldnumber)>8 then $fieldnumber=stringmid($fieldnumber, 1, 8)
	if (not fileexists($flimscreenpath)) or (_datetodayvalue($filemodifieddate[0], $filemodifieddate[1], $filemodifieddate[2])+7 < _datetodayvalue(@year, @mon, @mday)) then

		wait()
		mainframetype(2, 2, "flim{enter}")
		$screen=mainframegetscreen("FIELD INQUIRY MENU")
		mainframetype(2, 66, $fieldnumber)
		mainframetype(7, 4, "s{enter}")
		;send("{enter}")
		$screen=mainframegetscreen(" ")
		do
			$screen=mainframegetscreen(" ")
		until stringinstr($screen, "*** GENERAL FIELD INQUIRY ***") or stringinstr($screen, "NO DATA FOUND FOR THIS FIELD")
		if stringinstr($screen, "NO DATA FOUND FOR THIS FIELD") then
			msgbox(16, "", $fieldnumber & " is not a valid field number")
			exit
		endif
		$output=fileopen($flimscreenpath, 2)
		filewrite($output, $screen)
		fileclose($output)
	endif

	$input=fileopen($flimscreenpath, 0)
	$screen=fileread($input)
	fileclose($input)


	$fieldname=""
	if stringinstr($screen, "*** GENERAL FIELD INQUIRY ***") then
		$fieldname=mainframecopy($screen, 4, 23, 35)
		$h2s=mainframecopy($screen, 14, 29, 35)
	endif


	return $fieldname
endfunc


func getcol($sheetobj, $coltitle)
	$c=1
	$found=0
	$lastcol=$sheetobj.cells.specialcells(11).column
	$findobject=$sheetobj.rows(1).find($coltitle, default, default, default, default, 1)
	tooltip($findobject.column)
	return $findobject.column

endfunc


func uppercase($str)
	$newstr=""
	for $i=1 to stringlen($str)
		$char=stringmid($str, $i, 1)
		if asc($char)>=97 and asc($char<=122) then $char=chr(asc($char)-32)
		$newstr &= $char
	Next
	return $newstr


EndFunc

func shortcutfieldnames($shortcutname)
	$input=fileopen(@scriptdir & "\shortcut field names.txt", 0)
	$fieldnumber=""
	while 1
		$line=filereadline($input)
		if @error=-1 then exitloop
		$temp=stringsplit($line, "	")
		if $temp[1]=$shortcutname then $fieldnumber=$temp[2]
	wend
	return $fieldnumber
endfunc
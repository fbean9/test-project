
Attribute VB_Name = "OptionsDeltaTool"
Option Explicit

' ============================================================
'  OPTIONS DELTA TOOL  –  Main entry point
' ============================================================
Sub CalculateDelta()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Options Delta Tool")

    Dim rawInput As String
    rawInput = Trim(ws.Range("B2").Value)

    If rawInput = "" Then
        MsgBox "Please paste an options trade into cell B2.", vbExclamation, "No Input"
        Exit Sub
    End If

    ' Clear previous results
    Call ClearResults(ws)

    ' ── Parse ──────────────────────────────────────────────
    Dim trade As TradeInfo
    Dim parseErr As String
    parseErr = ParseTrade(rawInput, trade)

    If parseErr <> "" Then
        ws.Range("B26:G26").Interior.Color = RGB(255, 235, 235)
        ws.Range("B26").Value = "PARSE ERROR: " & parseErr
        ws.Range("B26").Font.Color = RGB(180, 0, 0)
        ws.Range("B26").Font.Bold = True
        ws.Range("B26").Font.Size = 11
        ws.Range("B26").Font.Italic = False
        Exit Sub
    End If

    ' ── Display parsed summary ─────────────────────────────
    Call DisplayParsed(ws, trade)

    ' ── Pull deltas from Bloomberg ─────────────────────────
    Call FetchAndDisplayDeltas(ws, trade)
End Sub

' ============================================================
'  DATA STRUCTURES
' ============================================================
Type LegInfo
    ticker      As String
    expiry      As String   ' MM/DD/YY  Bloomberg format
    strike      As Double
    optType     As String   ' "C" or "P"
    direction   As Integer  ' +1 buy, -1 sell (from market-maker perspective)
    ratio       As Integer  ' 1 or 2 (for 1x2)
    bbTicker    As String
    delta       As Double
End Type

Type TradeInfo
    rawTicker   As String
    expMonth    As String    ' display only (e.g. "may")
    expYear     As Integer   ' display only
    expDay      As Integer   ' 0 = standard monthly (3rd Fri); 1-31 = specific calendar day
    expWeekNum  As Integer   ' 0 = monthly; 1-5 = Nth Friday; -1 = generic "weekly" (nearest Fri)
    expIsToday  As Boolean   ' True = 0DTE
    expiry      As String    ' resolved Bloomberg date string MM/DD/YY
    quantity    As Long
    price       As Double
    structure   As String    ' single / call_spread / put_spread / 1x2 / butterfly / straddle / strangle
    buySell     As String    ' "bought" or "sold"  (from trader perspective)
    numLegs     As Integer
    legs(1 To 4) As LegInfo
    ambiguous   As Boolean
    altMsg      As String
End Type

' ============================================================
'  PARSER
' ============================================================
Function ParseTrade(s As String, ByRef t As TradeInfo) As String
    Dim lower As String
    lower = LCase(s)
    ' Normalize separators
    lower = Replace(lower, " - ", " ")
    lower = Replace(lower, "—", " ")
    lower = Replace(lower, Chr(8212), " ")
    lower = Replace(lower, "  ", " ")
    lower = Trim(lower)

    ' ── Quantity ──────────────────────────────────────────
    t.quantity = ExtractQuantity(lower)
    If t.quantity = 0 Then t.quantity = 1

    ' ── Direction (trader perspective) ────────────────────
    t.buySell = ExtractDirection(lower)
    If t.buySell = "" Then
        ParseTrade = "Could not determine bought/sold. Include 'bought', 'bot', 'sold', or 'sld'."
        Exit Function
    End If

    ' ── Ticker ────────────────────────────────────────────
    t.rawTicker = ExtractTicker(lower)
    If t.rawTicker = "" Then
        ParseTrade = "Could not identify ticker symbol."
        Exit Function
    End If

    ' ── Expiry ────────────────────────────────────────────
    Dim expM As String, expY As Integer, expD As Integer, expWk As Integer, expToday As Boolean
    If Not ExtractExpiry(lower, expM, expY, expD, expWk, expToday) Then
        ParseTrade = "Could not identify expiry. Include a month name, date (e.g. May 9), week (wk2), or 0dte."
        Exit Function
    End If
    t.expMonth   = expM
    t.expYear    = expY
    t.expDay     = expD
    t.expWeekNum = expWk
    t.expIsToday = expToday
    t.expiry     = FormatExpiry(expM, expY, expD, expWk, expToday)

    ' ── Price ─────────────────────────────────────────────
    t.price = ExtractPrice(lower)

    ' ── Strikes & structure ───────────────────────────────
    ' Pass expD and quantity so the fallback strike scan excludes them
    Dim strikes() As Double
    Dim nStrikes As Integer
    nStrikes = ExtractStrikes(lower, strikes, expD, t.quantity)

    If nStrikes = 0 Then
        ParseTrade = "Could not identify strike price(s)."
        Exit Function
    End If

    ' ── Determine structure ───────────────────────────────
    Dim isCall As Boolean, isPut As Boolean
    isCall = (InStr(lower, "call") > 0) Or (InStr(lower, " c ") > 0)
    isPut  = (InStr(lower, "put")  > 0) Or (InStr(lower, " p ") > 0)
    Dim isStraddle  As Boolean: isStraddle  = (InStr(lower, "straddle") > 0)
    Dim isStrangle  As Boolean: isStrangle  = (InStr(lower, "strangle") > 0)
    Dim is1x2       As Boolean: is1x2       = (InStr(lower, "1x2") > 0)
    Dim isFly       As Boolean: isFly       = (InStr(lower, "fly") > 0) Or (InStr(lower, "butterfly") > 0)
    Dim isSpread    As Boolean: isSpread    = (InStr(lower, "spread") > 0) Or (InStr(lower, " vs ") > 0)

    ' Build legs
    t.numLegs = 0

    If isStraddle Then
        t.structure = "straddle"
        If nStrikes < 1 Then
            ParseTrade = "Straddle requires at least one strike."
            Exit Function
        End If
        Dim sStrike As Double: sStrike = strikes(0)
        ' Leg1: call, Leg2: put
        ' Trader buys straddle → MM sold both → MM buys stock (call leg) and sells stock (put leg)
        t.numLegs = 2
        t.legs(1) = MakeLeg(t.rawTicker, t.expiry, sStrike, "C", t.buySell, 1)
        t.legs(2) = MakeLeg(t.rawTicker, t.expiry, sStrike, "P", t.buySell, 1)

    ElseIf isStrangle Then
        t.structure = "strangle"
        ' Try to extract explicit put/call strikes from text like "3 put 4 call"
        Dim putStrike As Double, callStrike As Double
        If ExtractStrangleStrikes(lower, putStrike, callStrike) Then
            t.numLegs = 2
            t.legs(1) = MakeLeg(t.rawTicker, t.expiry, callStrike, "C", t.buySell, 1)
            t.legs(2) = MakeLeg(t.rawTicker, t.expiry, putStrike,  "P", t.buySell, 1)
        ElseIf nStrikes >= 2 Then
            ' Lower = put, higher = call
            t.numLegs = 2
            t.legs(1) = MakeLeg(t.rawTicker, t.expiry, strikes(1), "C", t.buySell, 1)
            t.legs(2) = MakeLeg(t.rawTicker, t.expiry, strikes(0), "P", t.buySell, 1)
        Else
            ParseTrade = "Strangle requires two strikes. Specify e.g. '3 put 4 call'."
            Exit Function
        End If

    ElseIf isFly Then
        If nStrikes < 3 Then
            ParseTrade = "Butterfly requires three strikes (e.g. 3/3.5/4)."
            Exit Function
        End If
        t.structure = IIf(isPut, "put_butterfly", "call_butterfly")
        Dim oType As String: oType = IIf(isPut, "P", "C")
        ' Classic fly: long 1x lower, short 2x middle, long 1x upper
        ' If trader bought fly → MM sold fly → MM is short 1x lower, long 2x middle, short 1x upper
        t.numLegs = 3
        Dim flyDir As String: flyDir = IIf(t.buySell = "bought", "sold", "bought")
        t.legs(1) = MakeLeg(t.rawTicker, t.expiry, strikes(0), oType, flyDir, 1)
        t.legs(2) = MakeLeg(t.rawTicker, t.expiry, strikes(1), oType, t.buySell, 2)
        t.legs(3) = MakeLeg(t.rawTicker, t.expiry, strikes(2), oType, flyDir, 1)

    ElseIf is1x2 Then
        If nStrikes < 2 Then
            ParseTrade = "1x2 requires two strikes."
            Exit Function
        End If
        t.structure = IIf(isPut, "1x2_put", "1x2_call")
        Dim otType As String: otType = IIf(isPut, "P", "C")
        ' Trader bought 1x2: long 1x lower strike, short 2x upper strike
        t.numLegs = 2
        t.legs(1) = MakeLeg(t.rawTicker, t.expiry, strikes(0), otType, t.buySell, 1)
        Dim leg2Dir As String: leg2Dir = IIf(t.buySell = "bought", "sold", "bought")
        t.legs(2) = MakeLeg(t.rawTicker, t.expiry, strikes(1), otType, leg2Dir, 2)

    ElseIf isSpread Or nStrikes >= 2 Then
        If nStrikes < 2 Then
            ParseTrade = "Spread requires two strikes."
            Exit Function
        End If

        If isCall And isPut Then
            ' Ambiguous: could be call spread or put spread
            t.ambiguous = True
            t.altMsg = "Input mentions both 'call' and 'put'. Did you mean a call spread, put spread, or strangle?"
        End If

        If isPut And Not isCall Then
            t.structure = "put_spread"
            ' Put spread: higher strike = bought put, lower strike = sold put (or vice versa if sold)
            ' Convention: strikes(0) < strikes(1), higher strike put has larger abs delta
            ' Bought put spread: long higher strike put, short lower strike put
            t.numLegs = 2
            Dim psDir1 As String, psDir2 As String
            If t.buySell = "bought" Then
                psDir1 = "bought": psDir2 = "sold"
            Else
                psDir1 = "sold":   psDir2 = "bought"
            End If
            t.legs(1) = MakeLeg(t.rawTicker, t.expiry, strikes(1), "P", psDir1, 1)
            t.legs(2) = MakeLeg(t.rawTicker, t.expiry, strikes(0), "P", psDir2, 1)
        Else
            t.structure = "call_spread"
            ' Call spread bought: long lower strike call, short upper strike call
            Dim csDir1 As String, csDir2 As String
            If t.buySell = "bought" Then
                csDir1 = "bought": csDir2 = "sold"
            Else
                csDir1 = "sold":   csDir2 = "bought"
            End If
            t.numLegs = 2
            t.legs(1) = MakeLeg(t.rawTicker, t.expiry, strikes(0), "C", csDir1, 1)
            t.legs(2) = MakeLeg(t.rawTicker, t.expiry, strikes(1), "C", csDir2, 1)
        End If

    Else
        ' Single leg
        t.numLegs = 1
        t.structure = "single"
        Dim singType As String
        If isPut Then
            singType = "P"
        ElseIf isCall Then
            singType = "C"
        Else
            ' Try to infer from compact ticker like "may3.5C"
            If InStr(lower, "c" & Format(strikes(0), "0.##")) > 0 Or _
               InStr(lower, "c" & CStr(Int(strikes(0)))) > 0 Then
                singType = "C"
            ElseIf InStr(lower, "p" & Format(strikes(0), "0.##")) > 0 Or _
               InStr(lower, "p" & CStr(Int(strikes(0)))) > 0 Then
                singType = "P"
            Else
                ParseTrade = "Could not determine call or put."
                Exit Function
            End If
        End If
        t.legs(1) = MakeLeg(t.rawTicker, t.expiry, strikes(0), singType, t.buySell, 1)
    End If

    ParseTrade = ""   ' success
End Function

' ── Helper: build a LegInfo ───────────────────────────────────────────────────
Function MakeLeg(ticker As String, expiry As String, _
                 strike As Double, oType As String, _
                 traderDir As String, ratio As Integer) As LegInfo
    Dim lg As LegInfo
    lg.ticker   = UCase(ticker)
    lg.expiry   = expiry
    lg.strike   = strike
    lg.optType  = UCase(oType)
    lg.ratio    = ratio

    ' Market-maker perspective: MM is counterparty
    ' Trader buys call  → MM sold call  → MM needs to BUY stock (+delta)
    ' Trader sells call → MM bought call → MM needs to SELL stock (-delta)
    ' Trader buys put   → MM sold put   → MM needs to SELL stock (-delta)
    ' Trader sells put  → MM bought put → MM needs to BUY stock (+delta)
    If traderDir = "bought" Then
        If UCase(oType) = "C" Then
            lg.direction = 1   ' MM buys stock
        Else
            lg.direction = -1  ' MM sells stock
        End If
    Else  ' trader sold
        If UCase(oType) = "C" Then
            lg.direction = -1  ' MM sells stock
        Else
            lg.direction = 1   ' MM buys stock
        End If
    End If

    lg.bbTicker = BuildBBTicker(ticker, lg.expiry, oType, strike)
    MakeLeg = lg
End Function

' ── Bloomberg ticker: KEEL 5/3/24 C3.5 Equity ────────────────────────────────
Function BuildBBTicker(ticker As String, expiry As String, oType As String, _
                       strike As Double) As String
    Dim strikeStr As String
    If strike = Int(strike) Then
        strikeStr = Format(strike, "0")
    Else
        strikeStr = Format(strike, "0.##")
    End If
    BuildBBTicker = UCase(ticker) & " " & expiry & " " & UCase(oType) & strikeStr & " Equity"
End Function

' ── Format expiry → MM/DD/YY  (Bloomberg option ticker format) ───────────────
' Handles: 0DTE, specific day, Nth-Friday weekly, generic weekly, standard monthly
Function FormatExpiry(expM As String, expY As Integer, expD As Integer, _
                      expWk As Integer, expToday As Boolean) As String
    ' 0DTE: use today
    If expToday Then
        FormatExpiry = Format(Month(Now), "0") & "/" & Format(Day(Now), "0") & "/" & _
                       Format(Year(Now) Mod 100, "00")
        Exit Function
    End If

    Dim mNum As Integer: mNum = MonthNameToNum(expM)
    If mNum = 0 Then
        FormatExpiry = expM   ' already a formatted string or unrecognised — pass through
        Exit Function
    End If

    ' Specific calendar day (weekly/daily with known date)
    If expD > 0 Then
        FormatExpiry = Format(mNum, "0") & "/" & Format(expD, "0") & "/" & _
                       Format(expY Mod 100, "00")
        Exit Function
    End If

    ' Nth Friday of the month (wk1..wk5)
    If expWk > 0 Then
        Dim nthFri As Date: nthFri = NthFriday(expY, mNum, expWk)
        FormatExpiry = Format(Month(nthFri), "0") & "/" & Format(Day(nthFri), "0") & "/" & _
                       Format(Year(nthFri) Mod 100, "00")
        Exit Function
    End If

    ' Generic "weekly" (expWk = -1): nearest upcoming Friday in that month
    If expWk = -1 Then
        Dim d As Date: d = DateSerial(expY, mNum, 1)
        Dim endOfMonth As Date: endOfMonth = DateSerial(expY, mNum + 1, 0)
        Dim today As Date: today = Int(Now)
        Do While d <= endOfMonth
            If Weekday(d, vbSunday) = 6 Then   ' Friday
                If d >= today Then
                    FormatExpiry = Format(Month(d), "0") & "/" & Format(Day(d), "0") & "/" & _
                                   Format(Year(d) Mod 100, "00")
                    Exit Function
                End If
            End If
            d = d + 1
        Loop
        ' Fallback: first Friday of the month
        Dim ff As Date: ff = NthFriday(expY, mNum, 1)
        FormatExpiry = Format(Month(ff), "0") & "/" & Format(Day(ff), "0") & "/" & _
                       Format(Year(ff) Mod 100, "00")
        Exit Function
    End If

    ' Standard monthly: 3rd Friday
    Dim tf As Date: tf = NthFriday(expY, mNum, 3)
    FormatExpiry = Format(Month(tf), "0") & "/" & Format(Day(tf), "0") & "/" & _
                   Format(Year(tf) Mod 100, "00")
End Function

' Returns the Nth Friday of a given month/year (N=1 is first Friday)
Function NthFriday(yr As Integer, mo As Integer, n As Integer) As Date
    Dim d As Date: d = DateSerial(yr, mo, 1)
    Dim dow As Integer: dow = Weekday(d, vbSunday)  ' 1=Sun,2=Mon,...6=Fri,7=Sat
    Dim daysToFri As Integer: daysToFri = (6 - dow + 7) Mod 7
    NthFriday = d + daysToFri + (n - 1) * 7
End Function

Function MonthNameToNum(m As String) As Integer
    Select Case LCase(Left(m, 3))
        Case "jan": MonthNameToNum = 1
        Case "feb": MonthNameToNum = 2
        Case "mar": MonthNameToNum = 3
        Case "apr": MonthNameToNum = 4
        Case "may": MonthNameToNum = 5
        Case "jun": MonthNameToNum = 6
        Case "jul": MonthNameToNum = 7
        Case "aug": MonthNameToNum = 8
        Case "sep": MonthNameToNum = 9
        Case "oct": MonthNameToNum = 10
        Case "nov": MonthNameToNum = 11
        Case "dec": MonthNameToNum = 12
        Case Else:  MonthNameToNum = 0
    End Select
End Function

Function MonthNumToName(n As Integer) As String
    Dim names(12) As String
    names(1)="jan": names(2)="feb": names(3)="mar": names(4)="apr"
    names(5)="may": names(6)="jun": names(7)="jul": names(8)="aug"
    names(9)="sep": names(10)="oct": names(11)="nov": names(12)="dec"
    If n >= 1 And n <= 12 Then MonthNumToName = names(n) Else MonthNumToName = ""
End Function

Function NextYearForMonth(monthNum As Integer) As Integer
    Dim cy As Integer: cy = Year(Now)
    Dim cm As Integer: cm = Month(Now)
    If monthNum < cm Or (monthNum = cm And Day(Now) > 15) Then
        NextYearForMonth = cy + 1
    Else
        NextYearForMonth = cy
    End If
End Function

' ── Extract quantity (handles 5k, 5000x, 5000, 10m) ─────────────────────────
Function ExtractQuantity(s As String) As Long
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False

    ' Match patterns like 5k, 5000x, 5000, 10m
    re.Pattern = "\b(\d+(?:\.\d+)?)\s*(k|m)?\s*x?\b"
    Dim m As Object
    Dim matches As Object
    Set matches = re.Execute(s)

    Dim bestQty As Long
    bestQty = 0

    Dim i As Integer
    For i = 0 To matches.Count - 1
        Set m = matches(i)
        Dim num As Double
        num = CDbl(m.SubMatches(0))
        Dim suffix As String
        suffix = LCase(m.SubMatches(1))
        If suffix = "k" Then
            num = num * 1000
        ElseIf suffix = "m" Then
            num = num * 1000000
        End If
        If CLng(num) > bestQty And CLng(num) >= 100 Then
            bestQty = CLng(num)
        End If
    Next i

    ' Also check for patterns like "5000x" explicitly
    re.Pattern = "\b(\d{3,})\s*x\b"
    Set matches = re.Execute(s)
    For i = 0 To matches.Count - 1
        Dim q2 As Long
        q2 = CLng(matches(i).SubMatches(0))
        If q2 > bestQty Then bestQty = q2
    Next i

    ExtractQuantity = bestQty
End Function

' ── Extract direction ────────────────────────────────────────────────────────
Function ExtractDirection(s As String) As String
    If InStr(s, "bought") > 0 Or InStr(s, "bot") > 0 Or InStr(s, "buy") > 0 Then
        ExtractDirection = "bought"
    ElseIf InStr(s, "sold") > 0 Or InStr(s, "sld") > 0 Or InStr(s, "sell") > 0 Then
        ExtractDirection = "sold"
    Else
        ExtractDirection = ""
    End If
End Function

' ── Extract ticker: first all-uppercase or all-lower word that looks like a ticker
Function ExtractTicker(s As String) As String
    ' Try to find a word that is 1-6 letters appearing before expiry months
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True

    ' Look for word followed by month name or standalone uppercase block
    re.Pattern = "\b([a-z]{1,6})\s+(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)"
    Dim matches As Object
    Set matches = re.Execute(s)
    If matches.Count > 0 Then
        ExtractTicker = UCase(matches(0).SubMatches(0))
        Exit Function
    End If

    ' Fallback: find first word-like token that is not a keyword
    Dim skipWords As String
    skipWords = "|bought|bot|buy|sold|sld|sell|call|calls|put|puts|spread|fly|" & _
                "butterfly|straddle|strangle|1x2|for|at|the|color|vs|" & _
                "jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|"

    Dim parts() As String
    parts = Split(s, " ")
    Dim p As Integer
    For p = 0 To UBound(parts)
        Dim word As String
        word = Trim(parts(p))
        ' Remove trailing colon, $, etc.
        word = Replace(word, ":", "")
        If Len(word) >= 1 And Len(word) <= 6 Then
            If word Like "[a-zA-Z]*" And Not word Like "*[0-9]*" Then
                If InStr(skipWords, "|" & LCase(word) & "|") = 0 Then
                    ExtractTicker = UCase(word)
                    Exit Function
                End If
            End If
        End If
    Next p

    ExtractTicker = ""
End Function

' ── Extract expiry: handles 0DTE, M/D/YY, month+day, wkN, weekly, standard month ─
Function ExtractExpiry(s As String, ByRef expM As String, ByRef expY As Integer, _
                       ByRef expD As Integer, ByRef expWk As Integer, _
                       ByRef expToday As Boolean) As Boolean
    expD = 0: expWk = 0: expToday = False: expM = "": expY = Year(Now)

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True

    ' ── 0DTE / today ─────────────────────────────────────
    If InStr(s, "0dte") > 0 Or InStr(s, "0-dte") > 0 Or _
       InStr(s, "0 dte") > 0 Or InStr(s, "today") > 0 Then
        expToday = True
        expM = "today"
        ExtractExpiry = True
        Exit Function
    End If

    ' ── Explicit numeric date: M/D/YY or M/D/YYYY or M/D ─
    ' Only treat as a date if it does NOT look like a strike pattern (decimal/slash chain)
    ' Heuristic: both parts must be integers and the first must be 1-12
    re.Pattern = "\b(1[0-2]|[1-9])/([0-2]?\d|3[01])(?:/(\d{2,4}))?\b"
    Dim matches As Object
    Set matches = re.Execute(s)
    If matches.Count > 0 Then
        Dim dm As Object: Set dm = matches(0)
        Dim m1 As Integer: m1 = CInt(dm.SubMatches(0))
        Dim d1 As Integer: d1 = CInt(dm.SubMatches(1))
        If m1 >= 1 And m1 <= 12 And d1 >= 1 And d1 <= 31 Then
            expM  = MonthNumToName(m1)
            expD  = d1
            If dm.SubMatches(2) <> "" Then
                Dim yrx As Integer: yrx = CInt(dm.SubMatches(2))
                If yrx < 100 Then yrx = 2000 + yrx
                expY = yrx
            Else
                expY = NextYearForMonth(m1)
            End If
            ExtractExpiry = True
            Exit Function
        End If
    End If

    ' ── Month name ───────────────────────────────────────
    Dim months(11) As String
    months(0) = "jan": months(1) = "feb": months(2) = "mar"
    months(3) = "apr": months(4) = "may": months(5) = "jun"
    months(6) = "jul": months(7) = "aug": months(8) = "sep"
    months(9) = "oct": months(10) = "nov": months(11) = "dec"

    Dim i As Integer
    For i = 0 To 11
        If InStr(s, months(i)) > 0 Then
            expM = months(i)

            ' ── Year ────────────────────────────────────
            re.Global = False
            re.Pattern = months(i) & "\s*(\d{4})\b"
            Set matches = re.Execute(s)
            If matches.Count > 0 Then
                expY = CInt(matches(0).SubMatches(0))
            Else
                expY = NextYearForMonth(i + 1)
            End If
            re.Global = True

            ' ── Week indicator: wk1..wk5, week1..week5 ──
            re.Pattern = "(?:wk|week)\s*([1-5])\b"
            Set matches = re.Execute(s)
            If matches.Count > 0 Then
                expWk = CInt(matches(0).SubMatches(0))
                ExtractExpiry = True
                Exit Function
            End If

            ' ── Generic "weekly" with no week number ────
            If InStr(s, "weekly") > 0 Then
                expWk = -1
                ExtractExpiry = True
                Exit Function
            End If

            ' ── Specific day: "may 9", "may 9th", "may16"
            ' Pattern: month name immediately followed (optional space) by 1-2 digit day
            ' The day must be 1-31 and must NOT be followed by a decimal (to avoid "may 3.5")
            re.Pattern = months(i) & "\s*(\d{1,2})(?:st|nd|rd|th)?(?!\s*[\./\d])"
            Set matches = re.Execute(s)
            If matches.Count > 0 Then
                Dim dayCandidate As Integer
                dayCandidate = CInt(matches(0).SubMatches(0))
                If dayCandidate >= 1 And dayCandidate <= 31 Then
                    expD = dayCandidate
                End If
            End If

            ExtractExpiry = True
            Exit Function
        End If
    Next i
    ExtractExpiry = False
End Function

' ── Extract strikes from the string ─────────────────────────────────────────
' expDay: exclude this integer from fallback scan (it's the expiry day, not a strike)
' qty:    exclude this value from fallback scan (it's the quantity)
Function ExtractStrikes(s As String, ByRef strikes() As Double, _
                        expDay As Integer, qty As Long) As Integer
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True

    ' Look for slash-separated strikes: 3.5/4 or 3/3.5/4
    re.Pattern = "(\d+(?:\.\d+)?)\s*/\s*(\d+(?:\.\d+)?)\s*(?:/\s*(\d+(?:\.\d+)?))?"
    Dim matches As Object
    Set matches = re.Execute(s)
    If matches.Count > 0 Then
        Dim m As Object
        Set m = matches(0)
        Dim count As Integer: count = 0
        ReDim strikes(0 To 3)
        Dim v0 As Double: v0 = CDbl(m.SubMatches(0)): strikes(count) = v0: count = count + 1
        Dim v1 As Double: v1 = CDbl(m.SubMatches(1)): strikes(count) = v1: count = count + 1
        If m.SubMatches(2) <> "" Then
            strikes(count) = CDbl(m.SubMatches(2)): count = count + 1
        End If
        ' Sort ascending
        Call SortStrikes(strikes, count)
        ExtractStrikes = count
        Exit Function
    End If

    ' Look for straddle / strangle explicit strikes labeled "X put Y call" etc.
    re.Pattern = "(\d+(?:\.\d+)?)\s*(?:put|p)\b"
    Set matches = re.Execute(s)
    Dim putK As Double: putK = 0
    If matches.Count > 0 Then putK = CDbl(matches(0).SubMatches(0))

    re.Pattern = "(\d+(?:\.\d+)?)\s*(?:call|c)\b"
    Set matches = re.Execute(s)
    Dim callK As Double: callK = 0
    If matches.Count > 0 Then callK = CDbl(matches(0).SubMatches(0))

    If putK > 0 And callK > 0 Then
        ReDim strikes(0 To 1)
        strikes(0) = Application.Min(putK, callK)
        strikes(1) = Application.Max(putK, callK)
        ExtractStrikes = 2
        Exit Function
    End If

    ' Single strike: look for number after C/P or standalone decimal number
    ' Pattern: optional C or P immediately before/after a number
    re.Pattern = "[cp](\d+(?:\.\d+)?)\b"
    Set matches = re.Execute(s)
    If matches.Count > 0 Then
        ReDim strikes(0 To 0)
        strikes(0) = CDbl(matches(0).SubMatches(0))
        ExtractStrikes = 1
        Exit Function
    End If

    ' Fallback: look for decimal numbers or small integers that are plausible strikes
    ' Exclude: prices preceded by $, years, the expiry day, and the quantity
    re.Pattern = "(?<!\$)\b(\d+\.\d+|\d{1,4}(?!\d|k|m|,|\.))\b"
    Set matches = re.Execute(s)
    ReDim strikes(0 To 3)
    Dim cnt As Integer: cnt = 0
    Dim j As Integer
    For j = 0 To matches.Count - 1
        Dim v As Double: v = CDbl(matches(j).SubMatches(0))
        ' Must be a plausible strike: positive, not a year, not the expiry day, not the quantity
        Dim isYear    As Boolean: isYear    = (v >= 2020 And v <= 2040)
        Dim isExpDay  As Boolean: isExpDay  = (expDay > 0 And Abs(v - expDay) < 0.001)
        Dim isQty     As Boolean: isQty     = (qty > 0 And Abs(v - qty) < 0.001)
        If v >= 0.5 And Not isYear And Not isExpDay And Not isQty Then
            If cnt = 0 Or Abs(v - strikes(cnt - 1)) > 0.001 Then
                strikes(cnt) = v
                cnt = cnt + 1
                If cnt > 3 Then Exit For
            End If
        End If
    Next j
    Call SortStrikes(strikes, cnt)
    ExtractStrikes = cnt
End Function

Sub SortStrikes(ByRef arr() As Double, n As Integer)
    Dim i As Integer, j As Integer, tmp As Double
    For i = 0 To n - 2
        For j = 0 To n - i - 2
            If arr(j) > arr(j + 1) Then
                tmp = arr(j): arr(j) = arr(j + 1): arr(j + 1) = tmp
            End If
        Next j
    Next i
End Sub

' ── Extract price (first number after $ or "at" or "for") ───────────────────
Function ExtractPrice(s As String) As Double
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True
    re.Pattern = "\$(\d+(?:\.\d+)?)"
    Dim matches As Object
    Set matches = re.Execute(s)
    If matches.Count > 0 Then
        ExtractPrice = CDbl(matches(0).SubMatches(0))
        Exit Function
    End If
    re.Pattern = "(?:at|for)\s+(\d+(?:\.\d+)?)"
    Set matches = re.Execute(s)
    If matches.Count > 0 Then
        ExtractPrice = CDbl(matches(0).SubMatches(0))
        Exit Function
    End If
    ExtractPrice = 0
End Function

' ── Extract strangle put/call strikes from "3 put 4 call" ───────────────────
Function ExtractStrangleStrikes(s As String, ByRef putK As Double, ByRef callK As Double) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True

    re.Pattern = "(\d+(?:\.\d+)?)\s+put"
    Dim m1 As Object
    Set m1 = re.Execute(s)

    re.Pattern = "(\d+(?:\.\d+)?)\s+call"
    Dim m2 As Object
    Set m2 = re.Execute(s)

    If m1.Count > 0 And m2.Count > 0 Then
        putK  = CDbl(m1(0).SubMatches(0))
        callK = CDbl(m2(0).SubMatches(0))
        ExtractStrangleStrikes = True
    Else
        ExtractStrangleStrikes = False
    End If
End Function

' ============================================================
'  DISPLAY
' ============================================================
' Fixed layout rows — must match Python build script
' Row 7 : Parsed summary data    (ROW_PARSED)
' Row 8 : Ambiguity warning      (ROW_AMBIG)
' Row 10: Leg table headers      (ROW_LEGHDR)
' Rows 11-14: Leg data           (ROW_LEG1)
' Row 17: Delta section header
' Row 18: Delta table headers    (ROW_DELTAHDR)
' Rows 19-22: Delta data         (ROW_DELTA1)
' Row 26: Net result / error     (ROW_RESULT)
Const ROW_PARSED   As Long = 7
Const ROW_AMBIG    As Long = 8
Const ROW_LEG1     As Long = 11
Const ROW_DELTA1   As Long = 19
Const ROW_RESULT   As Long = 26

Sub ClearResults(ws As Worksheet)
    Dim r As Long
    ' Clear parsed row
    ws.Range("B7:G7").ClearContents
    ws.Range("B7:G7").Interior.Color = RGB(220, 230, 241)
    ws.Range("B7:G7").Font.Color = RGB(50, 50, 50)
    ' Clear ambiguity row
    ws.Range("B8:G8").ClearContents
    ws.Range("B8:G8").Interior.Color = RGB(255, 242, 204)
    ws.Range("B8").Font.Bold = False
    ws.Range("B8").Font.Color = RGB(0, 0, 0)
    ' Clear leg data rows 11-14
    ws.Range("B11:G14").ClearContents
    ws.Range("B11:G14").Font.Color = RGB(50, 50, 50)
    For r = 11 To 14
        ws.Range("B" & r & ":G" & r).Interior.Color = IIf(r Mod 2 = 1, RGB(255, 255, 255), RGB(220, 230, 241))
    Next r
    ' Clear delta data rows 19-22
    ws.Range("B19:G22").ClearContents
    ws.Range("B19:G22").Font.Color = RGB(50, 50, 50)
    For r = 19 To 22
        ws.Range("B" & r & ":G" & r).Interior.Color = IIf(r Mod 2 = 1, RGB(255, 255, 255), RGB(220, 230, 241))
    Next r
    ' Reset net result placeholder
    ws.Range("B26:G26").ClearContents
    ws.Range("B26:G26").Interior.Color = RGB(235, 243, 251)
    ws.Range("B26").Value = "NET DELTA IMPACT WILL APPEAR HERE AFTER CLICKING CALCULATE"
    ws.Range("B26").Font.Bold = True
    ws.Range("B26").Font.Size = 13
    ws.Range("B26").Font.Color = RGB(136, 136, 136)
    ws.Range("B26").Font.Italic = True
End Sub

Sub DisplayParsed(ws As Worksheet, t As TradeInfo)
    ' Row 7: parsed summary
    ws.Cells(ROW_PARSED, 2).Value = UCase(t.rawTicker)
    ' Show the resolved expiry date and its type label
    Dim expiryLabel As String
    If t.expIsToday Then
        expiryLabel = t.expiry & " (0DTE)"
    ElseIf t.expDay > 0 Then
        expiryLabel = t.expiry & " (specific date)"
    ElseIf t.expWeekNum > 0 Then
        expiryLabel = t.expiry & " (wk" & t.expWeekNum & " Friday)"
    ElseIf t.expWeekNum = -1 Then
        expiryLabel = t.expiry & " (next Friday)"
    Else
        expiryLabel = t.expiry & " (monthly)"
    End If
    ws.Cells(ROW_PARSED, 3).Value = expiryLabel
    ws.Cells(ROW_PARSED, 4).Value = t.structure
    ws.Cells(ROW_PARSED, 5).Value = UCase(t.buySell)
    ws.Cells(ROW_PARSED, 6).Value = Format(t.quantity, "#,##0")
    ws.Cells(ROW_PARSED, 7).Value = IIf(t.price > 0, "$" & Format(t.price, "0.##"), Chr(8212))
    ws.Range("B7:G7").Font.Bold = True

    ' Row 8: ambiguity warning
    If t.ambiguous Then
        ws.Cells(ROW_AMBIG, 2).Value = "  AMBIGUOUS: " & t.altMsg
        ws.Cells(ROW_AMBIG, 2).Font.Color = RGB(180, 90, 0)
        ws.Cells(ROW_AMBIG, 2).Font.Bold = True
    End If

    ' Rows 11-14: leg data (leg headers already in row 10 from sheet design)
    Dim i As Integer
    For i = 1 To t.numLegs
        Dim r As Long: r = ROW_LEG1 + (i - 1)
        ws.Cells(r, 2).Value = i
        ws.Cells(r, 3).Value = t.legs(i).ticker
        ws.Cells(r, 4).Value = t.legs(i).strike
        ws.Cells(r, 5).Value = t.legs(i).optType
        Dim dirLabel As String
        Select Case t.legs(i).direction
            Case 1:  dirLabel = "MM BUYS stock"
            Case -1: dirLabel = "MM SELLS stock"
            Case Else: dirLabel = Chr(8212)
        End Select
        ws.Cells(r, 6).Value = dirLabel
        ws.Cells(r, 6).Font.Color = IIf(t.legs(i).direction = 1, RGB(0, 112, 0), RGB(180, 0, 0))
        ws.Cells(r, 7).Value = t.legs(i).bbTicker
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Interior.Color = _
            IIf(i Mod 2 = 1, RGB(255, 255, 255), RGB(220, 230, 241))
    Next i
End Sub

Sub FetchAndDisplayDeltas(ws As Worksheet, t As TradeInfo)
    ' Delta table headers already in row 18 from sheet design
    ' Write delta data to rows 19-22
    Dim netDelta As Double: netDelta = 0
    Dim bbOk As Boolean: bbOk = True
    Dim i As Integer

    For i = 1 To t.numLegs
        Dim r As Long: r = ROW_DELTA1 + (i - 1)
        ws.Cells(r, 2).Value = i
        ws.Cells(r, 3).Value = t.legs(i).bbTicker

        Dim deltaVal As Variant
        On Error Resume Next
        deltaVal = Application.Run("BDP", t.legs(i).bbTicker, "DELTA_MID_RT")
        If Err.Number <> 0 Then deltaVal = CVErr(xlErrNA): bbOk = False
        Err.Clear
        On Error GoTo 0

        If IsError(deltaVal) Or IsEmpty(deltaVal) Or deltaVal = "" Then
            ws.Cells(r, 4).Value = "#N/A  Bloomberg unavailable or option not found"
            ws.Cells(r, 4).Font.Color = RGB(180, 0, 0)
            bbOk = False
        Else
            Dim dv As Double: dv = CDbl(deltaVal)
            t.legs(i).delta = Abs(dv)
            ws.Cells(r, 4).Value = Format(dv, "0.000")
            ws.Cells(r, 4).Font.Color = IIf(dv >= 0, RGB(0, 112, 0), RGB(180, 0, 0))

            Dim dir As Integer: dir = t.legs(i).direction
            Dim ratio As Integer: ratio = t.legs(i).ratio
            Dim contrib As Double
            contrib = t.quantity * Abs(dv) * 100 * ratio * dir
            netDelta = netDelta + contrib

            ws.Cells(r, 5).Value = IIf(dir = 1, "BUY", "SELL")
            ws.Cells(r, 5).Font.Color = IIf(dir = 1, RGB(0, 112, 0), RGB(180, 0, 0))
            ws.Cells(r, 6).Value = "1x" & ratio
            ws.Cells(r, 7).Value = Format(contrib, "#,##0")
            ws.Cells(r, 7).Font.Color = IIf(contrib >= 0, RGB(0, 112, 0), RGB(180, 0, 0))
        End If
        ws.Range(ws.Cells(r, 2), ws.Cells(r, 7)).Interior.Color = _
            IIf(i Mod 2 = 1, RGB(255, 255, 255), RGB(220, 230, 241))
    Next i

    ' Row 26: net result
    ws.Range("B26:G26").ClearContents
    ws.Range("B26:G26").Interior.Color = RGB(235, 243, 251)
    ws.Range("B26").Font.Italic = False
    ws.Range("B26").Font.Size = 13
    ws.Range("B26").Font.Bold = True

    If Not bbOk Then
        ws.Range("B26").Value = "NET DELTA: Cannot compute — Bloomberg data unavailable"
        ws.Range("B26").Font.Color = RGB(180, 0, 0)
        ws.Range("B26:G26").Interior.Color = RGB(255, 235, 235)
    ElseIf Abs(netDelta) <= 5000 Then
        ws.Range("B26").Value = "Minimal delta impact — near delta neutral"
        ws.Range("B26").Font.Color = RGB(100, 100, 100)
    ElseIf netDelta > 0 Then
        ws.Range("B26").Value = FormatShares(netDelta) & " to BUY"
        ws.Range("B26").Font.Color = RGB(0, 112, 0)
    Else
        ws.Range("B26").Value = FormatShares(netDelta) & " to SELL"
        ws.Range("B26").Font.Color = RGB(180, 0, 0)
    End If

    With ws.Range("B26:G26")
        .Borders(xlEdgeLeft).LineStyle   = xlContinuous: .Borders(xlEdgeLeft).Weight   = xlMedium
        .Borders(xlEdgeRight).LineStyle  = xlContinuous: .Borders(xlEdgeRight).Weight  = xlMedium
        .Borders(xlEdgeTop).LineStyle    = xlContinuous: .Borders(xlEdgeTop).Weight    = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

Function FormatShares(n As Double) As String
    Dim absN As Double: absN = Abs(n)
    If absN >= 1000000 Then
        FormatShares = Format(absN / 1000000, "0.#") & "m shares"
    ElseIf absN >= 1000 Then
        FormatShares = Format(absN / 1000, "0.#") & "k shares"
    Else
        FormatShares = Format(absN, "0") & " shares"
    End If
End Function

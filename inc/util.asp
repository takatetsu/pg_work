<%
' -----------------------------------------------------------------------------
' 時間小数点換算
' 時間を少数点表示に変換する。
' 分を時間に換算するときの端数処理は、10分未満は切り上げ、
' 10m＝0.2h、20m＝0.4h、30m＝0.5h、40m＝0.7h、50m＝0.9hとして丸める。
' -----------------------------------------------------------------------------
Function hhmm2Float(hhmm)
  Dim temp
  Dim upper
  Dim lower
  If (Len(hhmm) < 4) Then
    hhmm = Right("0000" & hhmm, 4)
  End If
  temp = Left(hhmm, 2) * 60 + Right(hhmm, 2)
  If ((temp mod 10) <> 0) Then
    temp = temp - (temp mod 10) + 10
  End If
  upper = Fix(temp / 60)
  Select Case temp - upper * 60
    Case 0
      lower = 0.0
    Case 10
      lower = 0.2
    Case 20
      lower = 0.4
    Case 30
      lower = 0.5
    Case 40
      lower = 0.7
    Case 50
      lower = 0.9
    Case Else
      lower = 0.0
  End Select
  hhmm2Float = upper + lower
End Function

' -----------------------------------------------------------------------------
' 小数点表示の時間を分に換算
' 時間を分に換算するときの端数処理は下記のとおり
' 0.2h=10m、0.4h=20m、0.5h=30m、0.7h=40m、0.9h=50m
' -----------------------------------------------------------------------------
Function floatTime2min(t)
    Dim upper
    Dim lower
    If (IsNumeric(Trim(t))) Then
        upper = Fix(t)
        lower = t - upper
        temp  = 0
        
        If lower = 0 Then
            temp = 0
        ElseIf lower <= 0.3  Then
            temp = 10
        ElseIf lower <= 0.45 Then
            temp = 20
        ElseIf lower <= 0.6  Then
            temp = 30
        ElseIf lower <= 0.8  Then
            temp = 40
        ElseIf lower <= 0.92 Then
            temp = 50
        End If
        floatTime2min = upper * 60 + temp
    Else
        floatTime2min = 0
    End if
End Function

' -----------------------------------------------------------------------------
' 分小数点換算
' 引数の分を少数点表示の時間に変換する。
' 分を時間に換算するときの端数処理は、10分未満は切り上げ、
' 10m＝0.2h、20m＝0.4h、30m＝0.5h、40m＝0.7h、50m＝0.9hとして丸める。
' -----------------------------------------------------------------------------
Function mm2Float(mm)
  Dim temp
  Dim upper
  Dim lower
  If ((mm mod 10) <> 0) Then
    temp = mm - (mm mod 10) + 10
  Else
    temp = mm
  End If
  upper = Fix(temp / 60)
  Select Case temp - upper * 60
    Case 0
      lower = 0.0
    Case 10
      lower = 0.2
    Case 20
      lower = 0.4
    Case 30
      lower = 0.5
    Case 40
      lower = 0.7
    Case 50
      lower = 0.9
    Case Else
      lower = 0.999
  End Select
  mm2Float = upper + lower
End Function

' -----------------------------------------------------------------------------
' 時間小数点換算
' 引数の分を少数点表示の日にちに変換する。
' 時間を日にちに換算するときの端数処理は切り上げ、
' 1時間＝0.2日、2時間＝0.3日、3時間＝0.4日、4時間＝0.5日、5時間＝0.7日、6時間＝0.8日、7時間＝0.9日、7時間40分＝1日とする。
' -----------------------------------------------------------------------------
Function mm2FloatDay(mm)
    '分を460分で割って、小数点1桁で表示する。小数点2桁以下は切上げ。
    tempFix = Fix(mm / 460 * 10)
    temp    = mm / 460 * 10
    If temp = tempFix Then
        mm2FloatDay = Fix(mm / 460 * 10) / 10
    Else
        mm2FloatDay = Fix(mm / 460 * 10) / 10 + 0.1
    End If
End Function

' -----------------------------------------------------------------------------
' 分時換算
' 引数を分の数値で受け取り、時間に換算した値を返す。
' 引数を時間に換算できないとき、ゼロバイト文字を返す。
' -----------------------------------------------------------------------------
Function min2Time(strMin)
    If (IsNumeric(Trim(strMin))) Then
        strHour   = Fix(Trim(strMin) / 60)
        strMinute = Trim(strMin) - (strHour * 60)
        If strHour >= 100 Then
            min2Time  = Right("000" & strHour, 3) & ":" & Right("00" & strMinute, 2)
        Else
            min2Time  = Right("00"  & strHour, 2) & ":" & Right("00" & strMinute, 2)
        End If
    Else
        min2Time  = ""
    End If
End Function

' -----------------------------------------------------------------------------
' 時分換算
' 引数を hh:mm 形式で受け取り、分換算した値を返す。
' 引数がゼロバイト文字の時は0で返します。
' -----------------------------------------------------------------------------
Function time2Min(strTime)
    If (Trim(strTime)="") Then
        time2Min = 0
    Else
        tempTime = editTime(strTime)
        time2Min = Left(tempTime, 2) * 60 + Right(tempTime, 2)
    End If
End Function

' -----------------------------------------------------------------------------
' 時刻フォーマット整形
' 引数をhh:mm形式の時刻に変換します。
' hhmmや数字を引数として受け取ったとき、前ゼロを付加し、hh:mmの形で返します。
' 時刻として妥当でないものは引数をそのまま返します。
' -----------------------------------------------------------------------------
Function editTime(strTime)
    Dim temp
    If (IsNumeric(Trim(Replace(strTime, ":", ""))) And Len(Trim(Replace(strTime, ":", ""))) <= 4) Then
        If (InStr(strTime, ":") > 0) Then
            temp = Right("00" & Left(Trim(strTime), InStr(Trim(strTime), ":") - 1), 2) & _
                   Right("00" & Right(Trim(strTime), Len(Trim(strTime)) - InStr(Trim(strTime), ":")), 2)
        Else
            temp = Right("0000" & Trim(strTime), 4)
        End If
        temp = Left(temp, 2) & ":" & Right(temp, 2)
    Else
        temp = Trim(strTime)
    End If

    Set objReTime           = new RegExp
    objReTime.Pattern       = "^([0-1][0-9]|[2][0-3]):[0-5][0-9]$"
    objReTime.IgnoreCase    = True
    objReTime.Global        = False
    If (objReTime.Test(temp)) Then
        ' 引数 strTime が hh:mm 形式の時
        editTime = temp
    Else
        ' 引数が時刻として妥当でない場合
        editTime = Trim(strTime)
    End If
End function

' -----------------------------------------------------------------------------
' 時刻妥当性チェック
' 引数をhh:mm形式の時刻として妥当か判定します。
' hhmmや数字を引数として受け取ったとき、前ゼロを付加し、hh:mmの形で検証します。
' ゼロバイト文字の場合は True を返します。
' -----------------------------------------------------------------------------
Function legalTime(strTime)
    If (Len(Trim(strTime)) = 0) Then
        legalTime = True
    Else
        Dim temp
        If (IsNumeric(Trim(Replace(strTime, ":", ""))) And Len(Trim(Replace(strTime, ":", ""))) <= 4) Then
            If (InStr(strTime, ":") > 0) Then
                temp = Right("00" & Left(Trim(strTime), InStr(Trim(strTime), ":") - 1), 2) & _
                       Right("00" & Right(Trim(strTime), Len(Trim(strTime)) - InStr(Trim(strTime), ":")), 2)
            Else
                temp = Right("0000" & Trim(strTime), 4)
            End If
            temp = Left(temp, 2) & ":" & Right(temp, 2)
        Else
            temp = Trim(strTime)
        End If

        Set objReTime           = new RegExp
        objReTime.Pattern       = "^([0-1][0-9]|[2][0-3]):[0-5][0-9]$"
        objReTime.IgnoreCase    = True
        objReTime.Global        = False
        If (objReTime.Test(temp)) Then
            legalTime = True
        Else
            legalTime = False
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' 経過分算出
' 引数1と引数2の時刻の差を分で返します。
' 引数はどちらも hh:mm の形式で渡してください。
' 引数が時刻として扱えないときは 0 を返します。
' time1 < time2 の時は24時間(1440分)加算して返します。
' -----------------------------------------------------------------------------
Function minDif(time1, time2)
    Set objReTime           = new RegExp
    objReTime.Pattern       = "^([0-1][0-9]|[2][0-3]):[0-5][0-9]$"
    objReTime.IgnoreCase    = True
    objReTime.Global        = False
    If (objReTime.Test(time1) And objReTime.Test(time2)) Then
        minDif = DateDiff("n", TimeValue(time1), TimeValue(time2))
        If (minDif < 0) Then
            minDif = 1440 + minDif
        End If
    Else
        minDif = 0
    End If
End Function

' -----------------------------------------------------------------------------
' 経過分算出
' 引数1と引数2の時刻の差を分で返します。
' 引数はどちらも hh:mm の形式で渡してください。
' 引数が時刻として扱えないときは 0 を返します。
' time1 < time2 の時はマイナス結果をプラスに変換して返します。
' -----------------------------------------------------------------------------
Function minDifIV(time1, time2)
    Set objReTime           = new RegExp
    objReTime.Pattern       = "^([0-1][0-9]|[2][0-3]):[0-5][0-9]$"
    objReTime.IgnoreCase    = True
    objReTime.Global        = False
    If (objReTime.Test(time1) And objReTime.Test(time2)) Then
        minDifIV = DateDiff("n", TimeValue(time1), TimeValue(time2))
        If (minDifIV < 0) Then
            minDifIV = minDifIV * -1
        End If
    Else
        minDifIV = 0
    End If
End Function

' -----------------------------------------------------------------------------
' 時間外深夜対象時間算出
' 引数1と引数2の間で、22:00~05:00までの時間外深夜の対象時間を分で返します。
' 引数が時刻として扱えないときは 0 を返します。
' -----------------------------------------------------------------------------
Function deepTimeMin(time1, time2)
    deepTimeMin = 0
    If (legalTime(time1) And legalTime(time2)) Then
        tempTime1 = editTime(time1)
        tempTime2 = editTime(time2)
        If ((tempTime1 >= "22:00" And tempTime2 > "22:00" And tempTime1 > tempTime2)  Or _
            (tempTime1 <  "05:00" And tempTime2 > "22:00"                          )) Then
            ' 深夜時間に2度かかる場合、再帰呼出しで2度集計処理を行う。
            deepTimeMin = deepTimeMin(tempTime1, "05:00") + deepTimeMin("22:00", tempTime2)
        Else
            ' 算出開始時刻設定
            startTime = ""
            If (time1 >= "22:00" Or time1 < "05:00") Then
                startTime = time1
            Else
                If (time2 > "22:00" Or time2 <= "05:00" Or time1 > time2) Then
                    startTime = "22:00"
                End If
            End If
            ' 算出終了時刻設定
            endTime = ""
            If (time2 >= "22:00" Or time2 < "05:00")  Then
                endTime = time2
            Else
                If (time1 > "22:00" Or time1 <= "05:00" Or time1 > time2) Then
                    endTime = "05:00"
                End If
            End If
            ' 時間外深夜業算出処理
            If (startTime = "" Or endTime = "") Then
                ' 時間外深夜対象外
                deepTimeMin = 0
            Else
                ' 時間外深夜対象
                deepTimeMin = DateDiff("n", TimeValue(startTime), TimeValue(endTime))
            End If
            If (deepTimeMin < 0) Then
                deepTimeMin = 1440 + deepTimeMin
            End If
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' 時系列チェック
' 引数で渡された時刻が t1 < t2 <= t3 < t4 の時系列かチェックし、
' 正常な時系列の場合は0を返す。エラー時は0以外を返す。
' 日付をまたぐ時系列も考慮するが、t1 ~ t4 までが24時間未満であること。
' -----------------------------------------------------------------------------
Function checkChronological(t1, t2, t3, t4)
    c1 = editTime(t1)
    c2 = editTime(t2)
    c3 = editTime(t3)
    c4 = editTime(t4)

    If (c1 < c2) Then
        If (c2 <= c3) Then
            If (c3 < c4) Then
                ' exp 08:30, 12:00, 13:00, 17:10
                checkChronological = 0
            Else
                If (c4 < c1) Then
                    ' exp 20:00, 22:00, 23:00, 03:00
                    checkChronological = 0
                Else
                    ' exp 20:00, 22:00, 23:00, 21:00
                    checkChronological = 1
                End If
            End If
        Else
            If (c3 < c4) Then
                If (c4 < c1) Then
                    ' exp 21:00, 23:00, 00:00, 05:00
                    checkChronological = 0
                Else
                    ' exp 21:00, 23:00, 02:00, 022:00
                    checkChronological = 2
                End If
            Else
                ' exp 21:00, 23:00, 02:00, 01:00
                checkChronological = 3
            End If
        End If
    Else
        If (c2 <= c3) Then
            If (c3 < c4) Then
                If (c4 < c1) Then
                    ' exp 23:00, 01:00, 02:00, 05:00
                    checkChronological = 0
                Else
                    ' exp 23:00, 01:00, 02:00, 23:30
                    checkChronological = 4
                End If
            Else
                ' exp 23:00, 03:00, 05:00, 04:00
                checkChronological = 5
            End If
        Else
            ' exp 23:00, 03:00, 02:00, (04:00)
            checkChronological = 6
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' 時系列チェック 未入力対応版
' 引数で渡された時刻が t1 < t2 <= t3 < t4 の時系列かチェックし、
' 正常な時系列の場合は0を返す。エラー時は0以外を返す。
' t1, t4 が未入力、もしくは t2, t3 が未入力の場合があるが、未入力の場合は正常な0を返す。
' 日付をまたぐ時系列も考慮するが、t1 ~ t4 までが24時間未満であること。
' -----------------------------------------------------------------------------
Function checkChronological_noentry_supported(t1, t2, t3, t4)
    c1 = editTime(t1)
    c2 = editTime(t2)
    c3 = editTime(t3)
    c4 = editTime(t4)

    If (Len(c1) = 0 And Len(c2) = 0) Or _
       (Len(c3) = 0 And Len(c4) = 0) Then
        checkChronological_noentry_supported = 0
    Else
        checkChronological_noentry_supported = checkChronological(t1, t2, t3, t4)
    End If
End Function


' -----------------------------------------------------------------------------
' 時系列チェック2
' 引数で渡された時刻が t1 <= t2 < t3 <= t4 の時系列かチェックし、
' 正常な時系列の場合は0を返す。エラー時は0以外を返す。
' 日付をまたぐ時系列も考慮するが、t1 ~ t4 までが24時間未満であること。
' -----------------------------------------------------------------------------
Function checkChronological2(t1, t2, t3, t4)
    c1 = editTime(t1)
    c2 = editTime(t2)
    c3 = editTime(t3)
    c4 = editTime(t4)

    If (c1 <= c2) Then
        If (c2 < c3) Then
            If (c3 <= c4) Then
                ' exp 08:30, 12:00, 13:00, 17:10
                checkChronological2 = 0
            Else
                If (c4 < c1) Then
                    ' exp 20:00, 22:00, 23:00, 03:00
                    checkChronological2 = 0
                Else
                    ' exp 20:00, 22:00, 23:00, 21:00
                    checkChronological2 = 1
                End If
            End If
        Else
            If (c3 <= c4) Then
                If (c4 < c1) Then
                    ' exp 21:00, 23:00, 00:00, 05:00
                    checkChronological2 = 0
                Else
                    ' exp 21:00, 23:00, 02:00, 022:00
                    checkChronological2 = 2
                End If
            Else
                ' exp 21:00, 23:00, 02:00, 01:00
                checkChronological2 = 3
            End If
        End If
    Else
        If (c2 < c3) Then
            If (c3 <= c4) Then
                If (c4 < c1) Then
                    ' exp 23:00, 01:00, 02:00, 05:00
                    checkChronological2 = 0
                Else
                    ' exp 23:00, 01:00, 02:00, 23:30
                    checkChronological2 = 4
                End If
            Else
                ' exp 23:00, 03:00, 05:00, 04:00
                checkChronological2 = 5
            End If
        Else
            ' exp 23:00, 03:00, 02:00, (04:00)
            checkChronological2 = 6
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' 時系列チェック2 未入力対応版
' 引数で渡された時刻が t1 <= t2 < t3 <= t4 の時系列かチェックし、
' 正常な時系列の場合は0を返す。エラー時は0以外を返す。
' t1, t4 が未入力、もしくは t2, t3 が未入力の場合があるが、未入力の場合は正常な0を返す。
' 日付をまたぐ時系列も考慮するが、t1 ~ t4 までが24時間未満であること。
' -----------------------------------------------------------------------------
Function checkChronological2_noentry_supported(t1, t2, t3, t4)
    c1 = editTime(t1)
    c2 = editTime(t2)
    c3 = editTime(t3)
    c4 = editTime(t4)

    If (Len(c1) = 0 And Len(c4) = 0) Or _
       (Len(c2) = 0 And Len(c3) = 0) Then
        checkChronological2_noentry_supported = 0
    Else
        checkChronological2_noentry_supported = checkChronological2(t1, t2, t3, t4)
    End If
End Function

' -----------------------------------------------------------------------------
' 時系列チェック3
' 引数で渡された時刻が t1 < t2 < t3 < t4 の時系列かチェックし、
' 正常な時系列の場合は0を返す。エラー時は0以外を返す。
' t1, t4 が未入力、もしくは t2, t3 が未入力の場合があるが、未入力の場合は正常な0を返す。
' 日付をまたぐ時系列も考慮するが、t1 ~ t4 または t3 ~ t2 までが24時間未満であること。
' -----------------------------------------------------------------------------
Function checkChronological3(t1, t2, t3, t4)
    c1 = editTime(t1)
    c2 = editTime(t2)
    c3 = editTime(t3)
    c4 = editTime(t4)

    If (Len(c1) = 0 And Len(c4) = 0) Or _
       (Len(c2) = 0 And Len(c3) = 0) Then
       checkChronological3 = 0
    Else
        If (c1 < c2) Then
            If (c2 < c3) Then
                If (c3 < c4) Then
                    ' exp 08:30, 12:00, 13:00, 17:10
                    checkChronological3 = 0
                Else
                    If (c4 < c1) Then
                        ' exp 20:00, 22:00, 23:00, 03:00
                        checkChronological3 = 0
                    Else
                        ' exp 20:00, 22:00, 23:00, 21:00
                        checkChronological3 = 1
                    End If
                End If
            Else
                If (c3 < c4) Then
                    If (c4 < c1) Then
                        ' exp 21:00, 23:00, 00:00, 05:00
                        checkChronological3 = 0
                    Else
                        ' exp 21:00, 23:00, 02:00, 022:00
                        checkChronological3 = 2
                    End If
                Else
                    ' exp 21:00, 23:00, 02:00, 01:00
                    checkChronological3 = 3
                End If
            End If
        Else
            If (c2 < c3) Then
                If (c3 < c4) Then
                    If (c4 < c1) Then
                        ' exp 23:00, 01:00, 02:00, 05:00
                        checkChronological3 = 0
                    Else
                        ' exp 23:00, 01:00, 02:00, 23:30
                        checkChronological3 = 4
                    End If
                Else
                    ' exp 23:00, 03:00, 05:00, 04:00
                    checkChronological3 = 5
                End If
            Else
                ' exp 23:00, 03:00, 02:00, (04:00)
                checkChronological3 = 6
            End If
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' 昼休みの時間にかかっているかチェック
' 引数で渡された時刻 t1～t2 が12:00～13:00の昼休みにかかっていれば、
' その時間（分）を返す。かかっていないとき、エラーのときは0を返す。
' -----------------------------------------------------------------------------
Function checkLunchTime(t1, t2)
    Dim tempBeginTime
    Dim tempEndTime
    checkLunchTime = 0
    If (legalTime(t1) And legalTime(t2)) Then
        If (t1 < "13:00" And t2 > "12:00") Then
            If t1 <= "12:00" Then
                tempBeginTime = "12:00"
            Else
                tempBeginTime = t1
            End if
            If t2 >= "13:00" Then
                tempEndTime   = "13:00"
            Else
                tempEndTime   = t2
            End if
            checkLunchTime    = minDif(tempBeginTime, tempEndTime)
        End If
    End If
End Function

' -----------------------------------------------------------------------------
' SQL Server TIMESTAMP 型を String 型に変換
' -----------------------------------------------------------------------------
Function TimestampToString(RsCol)
     Dim Buffer
     Dim i

     Buffer = "0x"
     'Buffer = ""
     For i = 1 To 8
         Buffer = Buffer & Right("00" & Hex(AscB(MidB(RsCol, i, 1))), 2)
     Next
     TimestampToString = Buffer
 End Function

' -----------------------------------------------------------------------------
' 引数1の時刻から引数2の時刻までの深夜時間帯と通常時間帯を時系列に配列にし、
' 分を返します。
' 例　　：引数1:04:00, 引数2:01:00
' 返り値：array(9,1,9,60,1020,180)
' 返り値は最初の3件までが区切った分の適用時間帯を表します。9が深夜時間帯、1が
' 通常時間帯です。4件目から6件目までが区切られた適用時間帯の分になります。
' 引数によっては、返り値の配列全てに値が入らない場合もあります。その場合は、
' 適用時間帯には空白が、時間の分は0が設定されます。
' -----------------------------------------------------------------------------
sub sepTime()
    sepTimeAry = array(0, 0, 0, 0, 0, 0)
    If     sepTime1 >= "05:00" And sepTime1 < "22:00" Then
        If sepTime2 >  "05:00" And sepTime2 <= "22:00" Then
            If sepTime1 < sepTime2 Then
                ' exp 08:00, 12:00 通常
                sepTimeAry(0) = 1
                sepTimeAry(3) = minDif(sepTime1, sepTime2)
            Else
                ' exp 20:00, 06:00 通常,深夜,通常
                sepTimeAry(0) = 1
                sepTimeAry(1) = 9
                sepTimeAry(2) = 1
                sepTimeAry(3) = minDif(sepTime1, "22:00" )
                sepTimeAry(4) = minDif("22:00" , "05:00" )
                sepTimeAry(5) = minDif("05:00" , sepTime2)
            End If
        ElseIf sepTime2 <= "05:00" Then
            ' exp 21:00, 02:00 通常,深夜
                sepTimeAry(0) = 1
                sepTimeAry(1) = 9
                sepTimeAry(3) = minDif(sepTime1, "22:00" )
                sepTimeAry(4) = minDif("22:00" , sepTime2)
        ElseIf sepTime2 >  "22:00" Then
            ' exp 11:00, 23:00 通常,深夜
                sepTimeAry(0) = 1
                sepTimeAry(1) = 9
                sepTimeAry(3) = minDif(sepTime1, "22:00" )
                sepTimeAry(4) = minDif("22:00" , sepTime2)
       End If
    ElseIf sepTime1 <  "05:00" Then
        If sepTime2 >  "05:00" And sepTime2 <= "22:00" Then
            ' exp 02:00, 08:00 深夜,通常
            sepTimeAry(0) = 9
            sepTimeAry(1) = 1
            sepTimeAry(3) = minDif(sepTime1, "05:00" )
            sepTimeAry(4) = minDif("05:00" , sepTime2)
        ElseIf sepTime2 <= "05:00" Then
            If sepTime1 < sepTime2 Then
                ' exp 01:00, 05:00 深夜
                sepTimeAry(0) = 9
                sepTimeAry(3) = minDif(sepTime1, sepTime2)
            Else
                ' exp 04:00, 02:00 深夜,通常,深夜
                sepTimeAry(0) = 9
                sepTimeAry(1) = 1
                sepTimeAry(2) = 9
                sepTimeAry(3) = minDif(sepTime1, "05:00" )
                sepTimeAry(4) = minDif("05:00" , "22:00" )
                sepTimeAry(5) = minDif("22:00" , sepTime2)
            End If
        ElseIf sepTime2 >  "22:00" Then
            ' exp 03:00, 23:30 深夜,通常,深夜
            sepTimeAry(0) = 9
            sepTimeAry(1) = 1
            sepTimeAry(2) = 9
            sepTimeAry(3) = minDif(sepTime1, "05:00" )
            sepTimeAry(4) = minDif("05:00" , "22:00" )
            sepTimeAry(5) = minDif("22:00" , sepTime2)
        End If
    ElseIf sepTime1 >= "22:00" Then
        If sepTime2 >  "05:00" And sepTime2 <= "22:00" Then
            ' exp 22:00, 08:00 深夜,通常
            sepTimeAry(0) = 9
            sepTimeAry(1) = 1
            sepTimeAry(3) = minDif(sepTime1, "05:00" )
            sepTimeAry(4) = minDif("05:00" , sepTime2)
        ElseIf sepTime2 <= "05:00" Then
            ' exp 22:00, 05:00 深夜
            sepTimeAry(0) = 9
            sepTimeAry(3) = minDif(sepTime1, sepTime2)
        ElseIf sepTime2 >  "22:00" Then
            If sepTime1 < sepTime2 Then
                ' exp 22:30, 23:30 深夜
                sepTimeAry(0) = 9
                sepTimeAry(3) = minDif(sepTime1, sepTime2)
            Else
                ' exp 23:30, 22:30 深夜,通常,深夜
                sepTimeAry(0) = 9
                sepTimeAry(1) = 1
                sepTimeAry(2) = 9
                sepTimeAry(3) = minDif(sepTime1, "05:00" )
                sepTimeAry(4) = minDif("05:00" , "22:00" )
                sepTimeAry(5) = minDif("22:00" , sepTime2)
            End If
        End If
    End If
End Sub

' -----------------------------------------------------------------------------
' 時間外等算出1
'
'
' -----------------------------------------------------------------------------
Sub compOverTime()
    v_overtime                  = 0     ' 時間外
    v_overtimelate              = 0     ' 時間外深夜業
    v_holidayshift              = 0     ' 休日出勤
    v_holidayshiftovertime      = 0     ' 休出時間外
    v_holidayshiftlate          = 0     ' 休出深夜業
    v_holidayshiftovertimelate  = 0     ' 休出時間外深夜業
    'v_flexovermin               = 0     ' フレックス勤務入力時間外(分)
    'Response.AppendToLog "@@-- log --"
    If (dayErrorFlag(i) <> "error") Then
        If workshift <> "9" Then
            compOverTimeDetail()
        Else
            ' フレックス勤務のとき
            'tempWork = minDif(editTime(v_overtime_begin), editTime(v_overtime_end))
            'tempRest = minDif(editTime(v_rest_begin    ), editTime(v_rest_end    ))
            'v_flexovermin = min2time(tempWork - tempRest)
        End If
    End If
    
    If IsNumeric(v_overtime) Then
        If v_overtime                  = 0 Then
            v_overtime                 = ""
        End If
    End If
    If IsNumeric(v_overtimelate) Then
        If v_overtimelate              = 0 Then
           v_overtimelate              = ""
        End If
    End If
    If IsNumeric(v_holidayshift) Then
        If v_holidayshift              = 0 Then
           v_holidayshift              = ""
        End If
    End If
    If IsNumeric(v_holidayshiftovertime) Then
        If v_holidayshiftovertime      = 0 Then
           v_holidayshiftovertime      = ""
        End If
    End If
    If IsNumeric(v_holidayshiftlate) Then
        If v_holidayshiftlate          = 0 Then
           v_holidayshiftlate          = ""
        End If
    End If
    If IsNumeric(v_holidayshiftovertimelate) Then
        If v_holidayshiftovertimelate  = 0 Then
           v_holidayshiftovertimelate  = ""
        End If
    End If

End Sub
' -----------------------------------------------------------------------------
' 時間外等算出2(詳細)
' -----------------------------------------------------------------------------
Sub compOverTimeDetail()
    tempWork         = minDif(editTime(v_overtime_begin ), editTime(v_overtime_end))
    tempRest         = minDif(editTime(v_rest_begin     ), editTime(v_rest_end    ))
    tempRestDeepTime = deepTimeMin(v_rest_begin    , v_rest_end    )
    tempDeepTime     = deepTimeMin(v_overtime_begin, v_overtime_end) - tempRestDeepTime
    'Response.AppendToLog "@@tempWork=" & tempWork
    'Response.AppendToLog "@@tempRest=" & tempRest
    'Response.AppendToLog "@@tempRestDeepTime=" & tempRestDeepTime
    'Response.AppendToLog "@@tempDeepTime=" & tempDeepTime
    If (Trim(v_morningholiday) <> "0" And Trim(v_afternoonholiday) <> "0"  And _
        Trim(v_morningwork   ) <> "1" And Trim(v_afternoonwork   ) <> "1"  And _
        Trim(v_morningwork   ) <> "5" And Trim(v_afternoonwork   ) <> "5"  And _
        Trim(v_morningholiday) <> "3" And Trim(v_afternoonholiday) <> "3") Then
        ' 休日
        ' 計算条件
        '    (休日区分（午前）及び（午後）に入力が有り
        '     かつ、出勤区分が振替出勤で無く
        '     休日区分午前、午後どちらも有給休暇でない場合)
        If v_overtime_begin = "" Then
        Else
            sepTimeAry1 = array(0, 0, 0, 0, 0, 0)
            sepTimeAry2 = array(0, 0, 0, 0, 0, 0)
            If tempRest = 0 Then
                ' 休憩時間なし
                ' 時間外開始時間-休憩開始時間までの時間を求める
                sepTime1 = editTime(v_overtime_begin)
                sepTime2 = editTime(v_overtime_end  )
                sepTime
                sepTimeAry1 = sepTimeAry
            Else
                ' 休憩時間あり
                sepTime1 = editTime(v_overtime_begin)
                sepTime2 = editTime(v_rest_begin    )
                sepTime
                sepTimeAry1 = sepTimeAry
                sepTime1 = editTime(v_rest_end      )
                sepTime2 = editTime(v_overtime_end  )
                sepTime
                sepTimeAry2 = sepTimeAry
            End If
            
            ' 時間外集計処理
            ' 休日出勤限度時間
            holidayshift_limit = 460
            For t=0 To 2 Step 1
                If     sepTimeAry1(t) = 1 Then
                    ' 通常時間帯のとき
                    If ((holidayshift_limit - sepTimeAry1(t+3)) >= 0) Then
                        ' 休日出勤に集計
                        v_holidayshift              = v_holidayshift             +  sepTimeAry1(t+3)
                        holidayshift_limit          = holidayshift_limit         -  sepTimeAry1(t+3)
                    Else
                        ' 休出時間外（時間外）のとき
                        ' 7時間40分までは休日出勤に集計し、あふれた時間を休出時間外に集計する。
                        v_holidayshift              = v_holidayshift             +  holidayshift_limit
                        v_holidayshiftovertime      = v_holidayshiftovertime     + (sepTimeAry1(t+3) - holidayshift_limit)
                        holidayshift_limit          = 0
                    End If
                ElseIf sepTimeAry1(t) = 9 Then
                    ' 深夜時間帯のとき
                    If ((holidayshift_limit - sepTimeAry1(t+3)) >= 0) Then
                        ' 休出時間内（時間外ではない）のとき
                        ' 休出深夜に集計
                        v_holidayshiftlate          = v_holidayshiftlate         +  sepTimeAry1(t+3)
                        holidayshift_limit          = holidayshift_limit         -  sepTimeAry1(t+3)
                    Else
                        ' 休出時間外（時間外）のとき
                        ' 7時間40分までは休出深夜に集計し、あふれた時間を休出時間外深夜に集計する。
                        v_holidayshiftlate          = v_holidayshiftlate         +  holidayshift_limit
                        v_holidayshiftovertimelate  = v_holidayshiftovertimelate + (sepTimeAry1(t+3) - holidayshift_limit)
                        holidayshift_limit          = 0
                    End If
                End If
            Next
            ' 休憩後時間外集計処理
            For t=0 To 2 Step 1
                If     sepTimeAry2(t) = 1 Then
                    ' 通常時間帯のとき
                    If holidayshift_limit - sepTimeAry2(t+3) >= 0 Then
                        ' 休日出勤に集計
                        v_holidayshift              = v_holidayshift             +  sepTimeAry2(t+3)
                        holidayshift_limit          = holidayshift_limit         -  sepTimeAry2(t+3)
                    Else
                        ' 休出時間外（時間外）のとき
                        ' 7時間40分までは休日出勤に集計し、あふれた時間を休出時間外に集計する。
                        v_holidayshift              = v_holidayshift             +  holidayshift_limit
                        v_holidayshiftovertime      = v_holidayshiftovertime     + (sepTimeAry2(t+3) - holidayshift_limit)
                        holidayshift_limit          = 0
                    End If
                ElseIf sepTimeAry2(t) = 9 Then
                    ' 深夜時間帯のとき
                    If holidayshift_limit - sepTimeAry2(t+3) >= 0 Then
                        ' 休出時間内（時間外ではない）のとき
                        ' 休出深夜に集計
                        v_holidayshiftlate          = v_holidayshiftlate         +  sepTimeAry2(t+3)
                        holidayshift_limit          = holidayshift_limit         -  sepTimeAry2(t+3)
                    Else
                        ' 休出時間外（時間外）のとき
                        ' 7時間40分までは休出深夜に集計し、あふれた時間を休出時間外深夜に集計する。
                        v_holidayshiftlate          = v_holidayshiftlate         +  holidayshift_limit
                        v_holidayshiftovertimelate  = v_holidayshiftovertimelate + (sepTimeAry2(t+3) - holidayshift_limit)
                        holidayshift_limit          = 0
                    End If
                End If
            Next
            If v_holidayshift             > 0 Then
                v_holidayshift             = min2time(v_holidayshift)
            End if
            If v_holidayshiftlate         > 0 Then
                v_holidayshiftlate         = min2time(v_holidayshiftlate)
            End if
            If v_holidayshiftovertime     > 0 Then
                v_holidayshiftovertime     = min2time(v_holidayshiftovertime)
            End if
            If v_holidayshiftovertimelate > 0 Then
                v_holidayshiftovertimelate = min2time(v_holidayshiftovertimelate)
            End if
        End If
    Else
        ' 出勤日
        ' 計算
        '    時間外時間-休憩時間の時間の内、
        '    22：00~翌朝5：00までの時間帯の時間数を「時間外深夜業」へ加算
        '    それ以外の時間帯の時間数を「時間外」へ加算
        tempRealWork = tempWork - (tempDeepTime + tempRestDeepTime) - (tempRest - tempRestDeepTime)
        If (tempRealWork > 0) Then
            v_overtime     = min2Time(tempRealWork)
        End If
        If (tempDeepTime > 0) Then
            v_overtimelate = min2Time(tempDeepTime)
        End if
    End If
End Sub

' -----------------------------------------------------------------------------
' 勤務表入力権限が無い場合は、画面遷移
' -----------------------------------------------------------------------------
Sub checkUser()
    If (Session("MM_is_input") = "1") Then

    ElseIf (Session("MM_is_superior") = "1") Then
      MM_redirectPage = "checklist.asp"
      Response.Redirect(MM_redirectPage)
    ElseIf (Session("MM_is_charge") = "1") Then
      MM_redirectPage = "inputall.asp"
      Response.Redirect(MM_redirectPage)
    End If
End Sub

' -----------------------------------------------------------------------------
' オペレータ交代勤務時の午前追加日数算出
' -----------------------------------------------------------------------------
Function operatorAddDays(v_operator)
    operatorAddDays = 0
    If (v_operator  = "1" Or _
        v_operator  = "2" Or _
        v_operator  = "3" Or _
        v_operator  = "5" Or _
        v_operator  = "6" Or _
        v_operator  = "E" Or _
        v_operator  = "F") Then
        ' 生産オペレータの甲番、乙番、日勤甲、見習(甲)、見習(乙)のときは0.5加算
        operatorAddDays = 0.5
    End If
    If v_operator = "4" Then
        ' 生産オペレータの生産会議乙のとき、1.0加算
        operatorAddDays = 1.0
    End If
End Function

' -----------------------------------------------------------------------------
' オペレータの可出勤日数計算
' 出勤日数は月間日数から休暇日数を引いていくことで求める。
' 交替勤務の場合は可出勤日数を加算し、交替勤務で振休の場合は減算する。
' 引数について
' mh:午前休日区分 morningholiday
' ah:午後休日区分 afternoonholiday
' mw:午前出勤区分 morningwork
' aw:午後出勤区分 afternoonwork
' op:オペレタ区分 operator
' -----------------------------------------------------------------------------
Function operatorWorkDay(mh, ah, mw, aw, op)
    operatorWorkDay = 0
    opday = operatorAddDays(op)
    ' 休暇(減算)
    ' 公休日
    If mh = "1" Then
        sumWorkDays = sumWorkDays - 0.5
    End if
    If ah = "1" Then
        sumWorkDays = sumWorkDays - 0.5
    End if
    ' 有給休暇
    If mh = "3" Then
        sumWorkDays = sumWorkDays + opday
    End if
    ' 振替休暇
    If mh = "2" Then
        sumWorkDays = sumWorkDays - 0.5 + opday
'        If op = "4" Then
'            ' 生産会議乙の場合
'            sumWorkDays = sumWorkDays - 0.5 - (opday * 2)
'        Else
'            sumWorkDays = sumWorkDays - 0.5 - (opday * 2)
'        End If
    End If
    If ah = "2" Then
        sumWorkDays = sumWorkDays - 0.5
    End if
    ' 出勤(加算)
    ' 出勤
    If (mw = "9") Then
        sumWorkDays = sumWorkDays + opday
    End If
    ' 振替出勤
    If (mw = "1" Or mw = "5") Then
        sumWorkDays = sumWorkDays + 0.5 + opday
    End If
    If (aw = "1" Or aw = "5") Then
        sumWorkDays = sumWorkDays + 0.5
    End If
    
End Function

' -----------------------------------------------------------------------------
' オペレータ交代勤務入力画面 inputop.asp での休暇表示文言取得
' -----------------------------------------------------------------------------
Function getHolidayText(holiday)
    Select Case holiday
        Case "2"
            getHolidayText = "振休"
        Case "3"
            getHolidayText = "有休"
        Case "4"
            getHolidayText = "代休"
        Case "5"
            getHolidayText = "特休"
        Case "6"
            getHolidayText = "保休"
        Case "7"
            getHolidayText = "欠勤"
    End Select
End Function
' -----------------------------------------------------------------------------
' オペレータ交代勤務入力画面 inputop.asp での休出表示文言取得
' -----------------------------------------------------------------------------
Function getWorkText(work)
    Select Case work
        Case "1"
            getWorkText = "振替出勤"
        Case "2"
            getWorkText = "休出"
        Case "3"
            getWorkText = "休出半日未満"
        Case "5"
            getWorkText = "振替出勤(出張)"
        Case "6"
            getWorkText = "休出(出張)"
    End Select
End Function

' -----------------------------------------------------------------------------
' オペレータ月間入力画面での交替勤務別勤務回数集計配列の指標を求める
' 引数 op:交替勤務区分, mw:午前出勤区分, mh:午前休暇区分, aw:午後出勤区分, ah:午後休暇区分
' 指標 0:甲番, 1:乙番, 2:常日, 3:日1, 4:日2, 5:日3, 6:他
' -----------------------------------------------------------------------------
Function getOpIdx(op, mw, mh, aw, ah)
    getOpIdx = 0
    Select Case op
        Case "0" ' 入力なしだが、常日として扱う
            getOpIdx = 2
        Case "1" ' 甲(有)
            If (((mw = "1" Or mw = "4" Or mw = "5" Or mw = "9") And aw = "0")  Or _
                ((aw = "1" Or mw = "4" Or aw = "5" Or aw = "9") And mw = "0")) Then
                getOpIdx = 6 ' 他へ集計
            Else
                getOpIdx = 0
            End If
        Case "2" ' 乙(有)
            If (((mw = "1" Or mw = "4" Or mw = "5" Or mw = "9") And aw = "0")  Or _
                ((aw = "1" Or mw = "4" Or aw = "5" Or aw = "9") And mw = "0")) Then
                getOpIdx = 6 ' 他へ集計
            Else
                getOpIdx = 1
            End If
        Case "A" ' 常日
            getOpIdx = 2
        Case "B" ' 日1xd
            getOpIdx = 3
        Case "C" ' 日2
            getOpIdx = 4
        Case "D" ' 日3
            getOpIdx = 5
        Case "E" ' 甲(無)
            If (((mw = "1" Or mw = "4" Or mw = "5" Or mw = "9") And aw = "0")  Or _
                ((aw = "1" Or mw = "4" Or aw = "5" Or aw = "9") And mw = "0")) Then
                getOpIdx = 6 ' 他へ集計
            Else
                idx = 0
            End If
        Case "F" ' 乙(無)
            If (((mw = "1" Or mw = "4" Or mw = "5" Or mw = "9") And aw = "0")  Or _
                ((aw = "1" Or mw = "4" Or aw = "5" Or aw = "9") And mw = "0")) Then
                getOpIdx = 6 ' 他へ集計
            Else
                getOpIdx = 1
            End If
    End Select
End Function

' -----------------------------------------------------------------------------
' オペレータ月間入力画面での交替勤務別勤務回数集計値算出
' 引数 op:交替勤務区分, mw:午前出勤区分, mh:午前休暇区分, aw:午後出勤区分, ah:午後休暇区分
' -----------------------------------------------------------------------------
Function getOpDay(op, mw, mh, aw, ah)
    getOpDay = 0.0
    If mw = "1" Or mw = "4" Or mw = "5" Or mw = "9" Then
        If (op = "1" Or op = "2" Or op = "E" Or op = "F") And _
           ((mw = "1" Or mw = "4" Or mw = "5" Or mw = "9") And aw = "0") Then
            ' 他項目へ1として集計
            getOpDay = getOpDay + 1.0
        Else
            getOpDay = getOpDay + 0.5
        End If
    End If
    If aw = "1" Or aw = "4" Or aw = "5" Or aw = "9" Then
        getOpDay = getOpDay + 0.5
    End If
    If mh = "2" Then
        If (op = "1" Or op = "2" Or op = "E" Or op = "F") And _
           ((mw = "1" Or mw = "4" Or mw = "5" Or mw = "9") And aw = "0") Then
            ' 他項目へ-1として集計
            ' マイナス集計必要？
            'getOpDay = getOpDay - 1.0
        Else
            ' マイナス集計必要？
            'getOpDay = getOpDay - 0.5
        End If
    End If
    If ah = "2" Then
        ' マイナス集計必要？
        'getOpDay = getOpDay - 0.5
    End If
End Function
%>

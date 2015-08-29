Imports System.IO
Imports System.Windows.Forms
Imports System.Text

Module NullBlankReplacer

    ' コンストラクタ
    ''' <summary>
    ''' コンストラクタ
    ''' </summary>
    ''' <remarks></remarks>
    Sub New()
        logBuilder = New StringBuilder()
    End Sub


    ' フィールド
    ''' <summary>
    ''' ログ
    ''' </summary>
    ''' <remarks></remarks>
    Private logBuilder As StringBuilder

    ''' <summary>
    ''' 移行 ID
    ''' </summary>
    ''' <remarks></remarks>
    Private migId As String

    ''' <summary>
    ''' テンポラリフォルダ
    ''' </summary>
    ''' <remarks></remarks>
    Private tempFolder As String

    ''' <summary>
    ''' データフォルダ
    ''' </summary>
    ''' <remarks></remarks>
    Private dataFolder As String

    ''' <summary>
    ''' ログフォルダ
    ''' </summary>
    ''' <remarks></remarks>
    Private errLogFolder As String

    ' メソッド
    ''' <summary>
    ''' エントリポイント
    ''' </summary>
    ''' <param name="cmdArgs">
    ''' 第一引数：移行 ID
    ''' 第二引数：テンポラリデータフォルダ
    ''' 第三引数：データフォルダ
    ''' 第四引数：エラーログフォルダ
    ''' </param>
    ''' <remarks></remarks>
    Sub Main(ByVal cmdArgs() As String)
        Debug.WriteLine(DateTime.Now)
        migId = cmdArgs(0)
        tempFolder = cmdArgs(1)
        dataFolder = cmdArgs(2)
        errLogFolder = cmdArgs(3)

        ' 引数チェック
        If Not Validate() Then
            OutputLog()
            Exit Sub
        End If

        ' データ変換
        Dim encoding = System.Text.Encoding.GetEncoding("Shift_JIS")
        Try
            Dim path As String = tempFolder & "\" & migId & ".txt"
            Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(path)
            Using sr As New StreamReader(path, encoding)
                Using newDataFile As New StreamWriter(dataFolder & "\" & fileName & ".txt", False, encoding) With {.AutoFlush = True}

                    Dim rownum As Integer = 0
                    While Not sr.EndOfStream
                        Dim value = sr.ReadLine

                        ' NULL 対策ヘッダ行を読み飛ばす
                        If rownum < 2 Then
                            rownum = rownum + 1
                            Continue While
                        End If

                        ' データの置換をする。
                        If String.IsNullOrEmpty(value) Then
                            Continue While
                        End If

                        ' ログデータ以降は読込対象外
                        If value.Contains("抽出件数") Then
                            Exit While
                        End If

                        ' 1行単位の文字列なので、改行を付与し、書き込み時に Write メソッドを使用する。
                        value = value & vbCrLf
                        value = ReplaceNuLLAndEmptyString(value)
                        newDataFile.Write(value)
                    End While
                End Using
            End Using
        Catch ex As Exception
            WriteLog(ex.Message)
            WriteLog(ex.StackTrace)
        Finally
            OutputLog()
            Debug.WriteLine(DateTime.Now)
        End Try
    End Sub

    ''' <summary>
    ''' 文字列 NULL → 空文字
    ''' 空文字 → NULL文字
    ''' に変換する。
    ''' </summary>
    ''' <param name="value">文字列</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ReplaceNuLLAndEmptyString(value As String) As String
        Dim str As String = value

        ' 空文字の置き換え。 
        ' 連続する空文字の置き換えに対応するため、2度置換する。
        str = str.Replace(vbTab & String.Empty & vbTab, vbTab & vbNullChar & vbTab)
        str = str.Replace(vbTab & String.Empty & vbTab, vbTab & vbNullChar & vbTab)
        ' 末尾の空文字の置換
        str = str.Replace(vbTab & String.Empty & vbCrLf, vbTab & vbNullChar & vbCrLf)

        ' NULL 文字の置き換え。
        ' 連続する NULL 文字の置き換えに対応するため、2度置換する。
        str = str.Replace(vbTab & "NULL" & vbTab, vbTab & String.Empty & vbTab)
        str = str.Replace(vbTab & "NULL" & vbTab, vbTab & String.Empty & vbTab)
        ' 末尾の NULL 文字の置換
        str = str.Replace(vbTab & "NULL" & vbCrLf, vbTab & String.Empty & vbCrLf)
        Return str
    End Function

    ''' <summary>
    ''' コマンドライン引数の検証を行なう。
    ''' すべて有効な場合は True
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Validate() As Boolean
        Dim isValid As Boolean = True
        Dim notExistFolder As String = "{0}が存在しません。"
        If String.IsNullOrEmpty(migId) Then
            WriteLog("移行 ID が空です。")
            isValid = False
        End If
        If Not Directory.Exists(tempFolder) Then
            WriteLog(String.Format(notExistFolder, "テンポラリフォルダ:" & tempFolder))
            isValid = False
        End If
        If Not Directory.Exists(dataFolder) Then
            WriteLog(String.Format(notExistFolder, "データフォルダ:" & dataFolder))
            isValid = False
        End If
        If Not Directory.Exists(errLogFolder) Then
            WriteLog(String.Format(notExistFolder, "ログフォルダ:" & errLogFolder))
            isValid = False
        End If

        Return isValid
    End Function

    ''' <summary>
    ''' ログメッセージを登録する。
    ''' </summary>
    ''' <param name="message"></param>
    ''' <remarks></remarks>
    Private Sub WriteLog(message As String)
        logBuilder.AppendLine(message)
    End Sub

    ''' <summary>
    ''' ログをファイルに書き込む
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OutputLog()
        If logBuilder.Length = 0 Then
            Exit Sub
        End If

        Using sw As New StreamWriter(errLogFolder & "\ExportTable.log")
            sw.Write(logBuilder.ToString())
        End Using
    End Sub

End Module


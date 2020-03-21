Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Web
Imports System.Reflection

Public Class Form1

    Private arrAllWord As New List(Of CWordModel)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        'バージョン情報
        Label4.Text = "Ver:" & My.Application.Info.Version.ToString

        '画面の文字を初期化
        Label2.Text = ""
        Label3.Text = ""

        arrAllWord.AddRange(fncReadWordList("あ"))
        arrAllWord.AddRange(fncReadWordList("い"))
        arrAllWord.AddRange(fncReadWordList("う"))
        arrAllWord.AddRange(fncReadWordList("え"))
        arrAllWord.AddRange(fncReadWordList("お"))
        arrAllWord.AddRange(fncReadWordList("か"))
        arrAllWord.AddRange(fncReadWordList("き"))
        arrAllWord.AddRange(fncReadWordList("く"))
        arrAllWord.AddRange(fncReadWordList("け"))
        arrAllWord.AddRange(fncReadWordList("こ"))
        arrAllWord.AddRange(fncReadWordList("さ"))
        arrAllWord.AddRange(fncReadWordList("し"))
        arrAllWord.AddRange(fncReadWordList("す"))
        arrAllWord.AddRange(fncReadWordList("せ"))
        arrAllWord.AddRange(fncReadWordList("そ"))
        arrAllWord.AddRange(fncReadWordList("た"))
        arrAllWord.AddRange(fncReadWordList("ち"))
        arrAllWord.AddRange(fncReadWordList("つ"))
        arrAllWord.AddRange(fncReadWordList("て"))
        arrAllWord.AddRange(fncReadWordList("と"))
        arrAllWord.AddRange(fncReadWordList("な"))
        arrAllWord.AddRange(fncReadWordList("に"))
        arrAllWord.AddRange(fncReadWordList("ぬ"))
        arrAllWord.AddRange(fncReadWordList("ね"))
        arrAllWord.AddRange(fncReadWordList("の"))
        arrAllWord.AddRange(fncReadWordList("は"))
        arrAllWord.AddRange(fncReadWordList("ひ"))
        arrAllWord.AddRange(fncReadWordList("ふ"))
        arrAllWord.AddRange(fncReadWordList("へ"))
        arrAllWord.AddRange(fncReadWordList("ほ"))
        arrAllWord.AddRange(fncReadWordList("ま"))
        arrAllWord.AddRange(fncReadWordList("み"))
        arrAllWord.AddRange(fncReadWordList("む"))
        arrAllWord.AddRange(fncReadWordList("め"))
        arrAllWord.AddRange(fncReadWordList("も"))
        arrAllWord.AddRange(fncReadWordList("や"))
        arrAllWord.AddRange(fncReadWordList("ゆ"))
        arrAllWord.AddRange(fncReadWordList("よ"))
        arrAllWord.AddRange(fncReadWordList("ら"))
        arrAllWord.AddRange(fncReadWordList("り"))
        arrAllWord.AddRange(fncReadWordList("る"))
        arrAllWord.AddRange(fncReadWordList("れ"))
        arrAllWord.AddRange(fncReadWordList("ろ"))
        arrAllWord.AddRange(fncReadWordList("わ"))
        arrAllWord.AddRange(fncReadWordList("を"))
        arrAllWord.AddRange(fncReadWordList("ん"))
        arrAllWord.AddRange(fncReadWordList("が"))
        arrAllWord.AddRange(fncReadWordList("ぎ"))
        arrAllWord.AddRange(fncReadWordList("ぐ"))
        arrAllWord.AddRange(fncReadWordList("げ"))
        arrAllWord.AddRange(fncReadWordList("ご"))
        arrAllWord.AddRange(fncReadWordList("ざ"))
        arrAllWord.AddRange(fncReadWordList("じ"))
        arrAllWord.AddRange(fncReadWordList("ず"))
        arrAllWord.AddRange(fncReadWordList("ぜ"))
        arrAllWord.AddRange(fncReadWordList("ぞ"))
        arrAllWord.AddRange(fncReadWordList("だ"))
        arrAllWord.AddRange(fncReadWordList("ぢ"))
        arrAllWord.AddRange(fncReadWordList("づ"))
        arrAllWord.AddRange(fncReadWordList("で"))
        arrAllWord.AddRange(fncReadWordList("ど"))
        arrAllWord.AddRange(fncReadWordList("ば"))
        arrAllWord.AddRange(fncReadWordList("び"))
        arrAllWord.AddRange(fncReadWordList("ぶ"))
        arrAllWord.AddRange(fncReadWordList("べ"))
        arrAllWord.AddRange(fncReadWordList("ぼ"))
        arrAllWord.AddRange(fncReadWordList("ぱ"))
        arrAllWord.AddRange(fncReadWordList("ぴ"))
        arrAllWord.AddRange(fncReadWordList("ぷ"))
        arrAllWord.AddRange(fncReadWordList("ぺ"))
        arrAllWord.AddRange(fncReadWordList("ぽ"))

    End Sub

    ''' <summary>
    ''' goo辞書の五十音リストから四字熟語一覧を取得
    ''' </summary>
    ''' <param name="sParamChar"></param>
    ''' <returns></returns>
    Private Function fncReadWordList(ByVal sParamChar As String) As List(Of CWordModel)

        'URLからHTMLデータを取得
        Dim param As String = HttpUtility.UrlEncode(sParamChar)
        Dim web As WebClient = New WebClient()
        web.Encoding = System.Text.Encoding.UTF8
        Dim html As String = web.DownloadString("https://dictionary.goo.ne.jp/idiom/index/" & param & "/")

        'HTMLDocumentクラスにHTMLをセット
        Dim doc As HtmlAgilityPack.HtmlDocument = New HtmlAgilityPack.HtmlDocument()
        doc.LoadHtml(html)

        '四字熟語一覧を取得
        Dim nodes1 As HtmlAgilityPack.HtmlNodeCollection =
            doc.DocumentNode.SelectNodes("/html/body/div[2]/div[2]/div/div[1]/div/section/div/div/div[2]/div[2]/ul/li/a/p[1]")

        If nodes1 Is Nothing Then
            Return New List(Of CWordModel)
        End If

        '説明一覧を取得
        Dim nodes2 As HtmlAgilityPack.HtmlNodeCollection =
            doc.DocumentNode.SelectNodes("/html/body/div[2]/div[2]/div/div[1]/div/section/div/div/div[2]/div[2]/ul/li/a/p[2]")

        If nodes2 Is Nothing Then
            Return New List(Of CWordModel)
        End If

        'リストに追加
        Dim arrWordList As New List(Of CWordModel)
        For i = 0 To nodes1.Count - 1

            Dim clsWordModel As New CWordModel

            clsWordModel.sWord = nodes1(i).InnerText
            clsWordModel.sExplain = nodes2(i).InnerText

            arrWordList.Add(clsWordModel)

        Next

        Return arrWordList

    End Function

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

        '乱数を生成
        Dim r As New Random
        Dim iIndexNO As Integer = r.Next(0, arrAllWord.Count - 1)

        Label2.Text = arrAllWord(iIndexNO).sWord
        Label3.Text = arrAllWord(iIndexNO).sExplain

    End Sub
End Class


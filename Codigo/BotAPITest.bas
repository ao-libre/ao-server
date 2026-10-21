Attribute VB_Name = "BotAPITest"
''
' BotAPITest.bas - Unit tests for BotAPI JSON parsing.
' Call RunBotAPITests from a debug build to verify JSON.bas integration.
''
Option Explicit

Private m_passed As Integer
Private m_failed As Integer

Private Sub Assert(ByVal condition As Boolean, ByVal testName As String)
    If condition Then
        m_passed = m_passed + 1
        Debug.Print "[PASS] " & testName
    Else
        m_failed = m_failed + 1
        Debug.Print "[FAIL] " & testName
    End If
End Sub

Public Sub RunBotAPITests()
    m_passed = 0
    m_failed = 0

    Debug.Print "=== BotAPI JSON Tests ==="

    ' Test 1: parse action field
    Dim json1 As String
    json1 = "{""action"":""getState"",""charname"":""Bot1""}"
    Dim p1 As Object
    Set p1 = parse(json1)
    Call Assert(Not p1 Is Nothing, "parse returns non-null")
    Call Assert(p1.Exists("action"), "action key exists")
    Call Assert(CStr(p1("action")) = "getState", "action = getState")
    Call Assert(p1.Exists("charname"), "charname key exists")
    Call Assert(CStr(p1("charname")) = "Bot1", "charname = Bot1")

    ' Test 2: parse walk with heading
    Dim json2 As String
    json2 = "{""action"":""walk"",""charname"":""Bot1"",""heading"":""north""}"
    Dim p2 As Object
    Set p2 = parse(json2)
    Call Assert(Not p2 Is Nothing, "walk parse non-null")
    Call Assert(CStr(p2("action")) = "walk", "walk action")
    Call Assert(CStr(p2("heading")) = "north", "heading = north")

    ' Test 3: parse attack (no extra params)
    Dim json3 As String
    json3 = "{""action"":""attack"",""charname"":""TestBot""}"
    Dim p3 As Object
    Set p3 = parse(json3)
    Call Assert(CStr(p3("action")) = "attack", "attack action")
    Call Assert(CStr(p3("charname")) = "TestBot", "charname = TestBot")

    ' Test 4: parse castSpell with numeric slot
    Dim json4 As String
    json4 = "{""action"":""castSpell"",""charname"":""Bot1"",""slot"":3}"
    Dim p4 As Object
    Set p4 = parse(json4)
    Call Assert(CStr(p4("action")) = "castSpell", "castSpell action")
    Call Assert(CInt(p4("slot")) = 3, "slot = 3")

    ' Test 5: parse error handling - invalid JSON
    Dim p5 As Object
    Set p5 = parse("not valid json")
    ' JSON.bas returns Nothing on parse failure (non-object/array input sets m_parserrors
    ' but never assigns the return value, so parse returns the default Nothing)
    Call Assert(p5 Is Nothing, "invalid JSON returns Nothing")

    ' Test 6: BotDispatch getState returns error for unknown charname
    Dim resp As String
    resp = BotDispatch("{""action"":""getState"",""charname"":""__NO_SUCH_BOT__""}")
    Call Assert(InStr(resp, "false") > 0, "getState unknown char returns error")

    ' Test 7: BotDispatch unknown action returns error
    Dim resp2 As String
    resp2 = BotDispatch("{""action"":""teleport"",""charname"":""Bot1""}")
    Call Assert(InStr(resp2, "false") > 0, "unknown action returns error")

    ' Test 8: BotDispatch malformed JSON returns error
    Dim resp3 As String
    resp3 = BotDispatch("not json at all")
    Call Assert(InStr(resp3, "false") > 0, "malformed JSON returns error")

    Debug.Print "=== Results: " & m_passed & " passed, " & m_failed & " failed ==="
End Sub

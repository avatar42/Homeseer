' Based on code found in https://forums.homeseer.com/showthread.php?t=175817&styleid=8
' Lists out Event info so you can find events using by searching the generated spreadsheet. For instance to find all the events using a script of all the ones that speak which you can not do in the web interface. 

Const LT As String = "EventList"


Sub Main(Parm As Object)
    Dim file As System.IO.StreamWriter

    Try
        Dim EvListGroup() As strEventGroupData
        EvListGroup = hs.Event_Group_Info_All
        hs.writelog(LT, "Found " & EvListGroup.Length & " Events in list")
        file = My.Computer.FileSystem.OpenTextFileWriter("./html/reports/Event.list." & Now.ToString("yyMMdd") & ".csv", True)
        file.WriteLine("Group Name,Group ID,Event Name,Reference,Type,Last Triggered,A.T.GA.GC,Name")
        For Each EventGroup As strEventGroupData In EvListGroup
            hs.writelog(LT, "*** New Group ***")
            hs.writelog(LT, "Event Group Name: " & EventGroup.GroupName & " Group ID: " & EventGroup.GroupID)

            Dim EvListEvents() As strEventData

            hs.writelog(LT, "--> Process Events In Group")

            EvListEvents = hs.Event_Info_Group(EventGroup.GroupID)

            For Each SingleEvent As strEventData In EvListEvents
                hs.writelog(LT, "Event Name: " & SingleEvent.Event_Name & " Reference: " & SingleEvent.Event_Ref & " Type: " & SingleEvent.Event_Type)
                hs.writelog(LT, "This Event Was Last Triggered On: " & SingleEvent.Last_Triggered.ToString)

                hs.writelog(LT, "This Event Has " & SingleEvent.Action_Count & " Action(s)")

                If SingleEvent.Action_Count > 0 Then
                    Dim ActionList() As String = SingleEvent.Actions

                    For Each ActionName As String In ActionList
                        hs.writelog(LT, "Action Name: " & ActionName)
                        file.WriteLine(EventGroup.GroupName & "," & EventGroup.GroupID & "," & SingleEvent.Event_Name & "," & SingleEvent.Event_Ref & "," & SingleEvent.Event_Type & "," & SingleEvent.Last_Triggered.ToString & ",A," & ActionName)
                    Next
                End If

                hs.writelog(LT, "This Event Has " & SingleEvent.Trigger_Count & " Trigger(s)")

                hs.writelog(LT, "This Event Has " & SingleEvent.Trigger_Group_Count & " Group Trigger(s)")

                If SingleEvent.Trigger_Count > 0 Then
                    Dim TriggerList() As strEventTriggerGroupData = SingleEvent.Trigger_Groups
                    For Each SingleTrigger As strEventTriggerGroupData In TriggerList
                        Dim FinalTrigArray() As String = SingleTrigger.Triggers
                        For Each FinalTriggerName As String In FinalTrigArray
                            hs.writelog(LT, "Single Trigger Name: " & FinalTriggerName)
                            file.WriteLine(EventGroup.GroupName & "," & EventGroup.GroupID & "," & SingleEvent.Event_Name & "," & SingleEvent.Event_Ref & "," & SingleEvent.Event_Type & "," & SingleEvent.Last_Triggered.ToString & ",T," & FinalTriggerName)
                        Next
                    Next
                End If
            Next

            hs.writelog(LT, "--> Finished Event Processing")

            hs.writelog(LT, "This Event Group Has " & EventGroup.Global_Actions_Count & " Global Actions")

            If EventGroup.Global_Actions_Count > 0 Then
                Dim GlActionArr() As String = EventGroup.Global_Actions
                For Each GlAction As String In GlActionArr
                    hs.writelog(LT, "Action Name: " & GlAction)
                    file.WriteLine(EventGroup.GroupName & "," & EventGroup.GroupID & ",,,,,GA," & GlAction)
                Next
            End If

            hs.writelog(LT, "This Event Group Has " & EventGroup.Global_Conditions_Count & " Global Conditions")

            If EventGroup.Global_Conditions_Count > 0 Then
                Dim GlConditionArr() As String = EventGroup.Global_Conditions
                For Each GlConditionName As String In GlConditionArr
                    hs.writelog(LT, "Condition Name: " & GlConditionName)
                    file.WriteLine(EventGroup.GroupName & "," & EventGroup.GroupID & ",,,,,GC," & GlConditionName)
                Next
            End If
            hs.writelog(LT, "*** End Of Group ***")
        Next

    Catch ex As Exception : hs.writelog(LT, "Exception: " & ex.message)
    Finally
        file.Close()
    End Try
    hs.writelog(LT, "*** End Of Event List***")

End Sub
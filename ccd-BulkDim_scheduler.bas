Option Explicit
Public alertTimeOne As Double           '>>>>>>timers group<<<<<<<
Public alertTimeTwo As Double
Public alertTimeThree As Double
Public alertTimeFour As Double
'----------------------------------------------------------------------------------------------------
Public alertTimeOneAsNumber As Double      '>>>>>>timers used to check upper limits group<<<<<<
Public alertTimeTwoAsNumber As Double
Public alertTimeThreeAsNumber As Double
'---------------------------------------------------------------------------------------------------
Public customTimerTest As String        '>>>>>>>mgmt dashboard variable group<<<<<<<
Public defaultDash As String
Public defDashSched As Double
Public dashAdHoc As Boolean
Public restoreDashSched As Double
Public defDashSchedUpper As Double
Public restoreDashSchedUpper As Double
Public testDefDash As Double
Public testUpDefDash As Double
'---------------------------------------------------------------------------------------------------
Public telTimerTest As String           '>>>>>>>telogis variable group<<<<<<<
Public defaultTel As String
Public defTelSched As Double
Public defTelSchedUpper As Double
Public restoreTelSched As Double
Public restoreTelSchedUpper As Double
Public telAdHoc As Boolean
'---------------------------------------------------------------------------------------------------
Public intConTimerTest As String           '>>>>>>>internal consolidated variable group<<<<<<<
Public defaultIntCon As String
Public defIntConSched As Double
Public defIntConSchedUpper As Double
Public restoreIntConSched As Double
Public restoreIntConSchedUpper As Double
Public intConAdHoc As Boolean
'--------------------------------------------------------------------------------------------------
Public extLateTimerTest As String           '>>>>>>>external late variable group<<<<<<<
Public defaultExtLate As String
Public defExtLateSched As Double
Public defExtLateSchedUpper As Double
Public restoreExtLateSched As Double
Public restoreExtLateSchedUpper As Double
Public extLateAdHoc As Boolean
'-------------------------------------------------------------------------------------------------
Public myLateTrip_allroute As Variant                  '>>>>>>>>all route sub globals<<<<<<<<<<<<
Public myArr_allroute As Variant
Public myEndArr_allroute As Variant
Public lateTrip_allroute As Variant
Public myFranArr_allroute As Variant
Public myFranEndArr_allroute As Variant
Public myFranLateTrip_allroute As Variant
Public mainRecip_allroute As Variant
'------------------------------------------------------------------------------------------------
Public myLateTrip_alltrip As Variant                  '>>>>>>>>all trip sub globals<<<<<<<<<<<<
Public myArr_alltrip As Variant
Public myEndArr_alltrip As Variant
Public lateTrip_alltrip As Variant
Public myFranArr_alltrip As Variant
Public myFranEndArr_alltrip As Variant
Public myFranLateTrip_alltrip As Variant
Public mainRecip_alltrip As Variant
'-------------------------------------------------------------------------------------------------
Public myLateTrip_newdynamic As Variant                 '>>>>>>>>new dynamic sub globals<<<<<<<<<<<<
Public myArr_newdynamic As Variant
Public myEndArr_newdynamic As Variant
Public lateTrip_newdynamic As Variant
Public myFranArr_newdynamic As Variant
Public myFranEndArr_newdynamic As Variant
Public myFranLateTrip_newdynamic As Variant
Public mainRecip_newdynamic As Variant
Public myCustArr_newdynamic As Variant
Public myCustEndArr_newdynamic As Variant
Public myCustLateTrip_newdynamic As Variant
Public whenLateCheck As Integer
'---------------------------------------------------------------------------------------------------
Public myLateTrip_uiroute As Variant                 '>>>>>>>>user input route blast sub globals<<<<<<<<<<<<
Public myArr_uiroute As Variant
Public myEndArr_uiroute As Variant
Public lateTrip_uiroute As Variant
Public myFranArr_uiroute As Variant
Public myFranEndArr_uiroute As Variant
Public myFranLateTrip_uiroute As Variant
Public mainRecip_uiroute As Variant
Public myCustArr_uiroute As Variant
Public myCustEndArr_uiroute As Variant
Public myCustLateTrip_uiroute As Variant
'---------------------------------------------------------------------------------------------------
Public myLateTrip_uidynamic As Variant                 '>>>>>>>>user input dynamic sub globals<<<<<<<<<<<<
Public myArr_uidynamic As Variant
Public myEndArr_uidynamic As Variant
Public lateTrip_uidynamic As Variant
Public myFranArr_uidynamic As Variant
Public myFranEndArr_uidynamic As Variant
Public myFranLateTrip_uidynamic As Variant
Public mainRecip_uidynamic As Variant
Public myCustArr_uidynamic As Variant
Public myCustEndArr_uidynamic As Variant
Public myCustLateTrip_uidynamic As Variant

Public Sub declareAlerts()
'alert timers listed here

alertTimeOne = 0
alertTimeTwo = 0
alertTimeThree = 0
alertTimeFour = 0

'>>>>>>>>>>management dashboard cycle timer variables listed here<<<<<<<<<<<'

customTimerTest = ""  'this is used to define a custom timer for mgmt dash event
defaultDash = "04:00:00"   'this is the default cycle var for mgmt dash event
defDashSched = TimeValue("12:05:00")   'this is the default earliest time var for mgmt dash schedule
defDashSchedUpper = TimeValue("16:15:00")    'this is the default latest time var for mgmt dash schedule
restoreDashSched = TimeValue("12:05:00")   'this the var that restores the default earlies time for mgmt dash schedule
restoreDashSchedUpper = TimeValue("16:15:00")  'this is the var that restores the default latest time for mgmt dash schedule
dashAdHoc = False  'this is the mgmt dashboard specific adHoc indicator - there is logic predicated on this var
testDefDash = TimeValue("18:10:00")   'these are test vars that allow me to show a schedule cycle can be started and stopped and started again
testUpDefDash = TimeValue("18:16:00") 'same as above - this is the upper version of the above lower version

'>>>>>>>>>>get live telogis cycle timer variables listed here<<<<<<<<<<<' 'follows same logic as above timers and descriptions

telTimerTest = ""
defaultTel = "01:00:00"
defTelSched = TimeValue("06:00:00")
defTelSchedUpper = TimeValue("16:10:00")
restoreTelSched = TimeValue("06:00:00")
restoreTelSchedUpper = TimeValue("16:10:00")
telAdHoc = False


'>>>>>>>>>>internal consolidated cycle timer variables listed here<<<<<<<<<<<' 'follows same logic as above timers and descriptions

intConTimerTest = ""
defaultIntCon = "01:00:00"
defIntConSched = TimeValue("06:10:00")
defIntConSchedUpper = TimeValue("16:15:00")
restoreIntConSched = TimeValue("06:10:00")
restoreIntConSchedUpper = TimeValue("16:15:00")
intConAdHoc = False

End Sub


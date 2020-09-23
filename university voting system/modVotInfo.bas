Attribute VB_Name = "Module1"
Public MyPicture As String
Public ToPass As Integer

Type Positions
    Position  As String * 200
End Type


Type Parties
    Party  As String * 200
End Type


Type VoterInfo
    VLastName   As String * 60
    VFirstName  As String * 60
    VMidName   As String * 60
    VStudNum   As String * 60
    VCourse    As String * 200
    VValidator As String * 500
    VFullName  As String * 60
    VTimeTrack As String * 60
    VVoteTimeTrack As String * 60
    VVote      As String * 10000
    VVoteTrack As Integer
    
    

    
End Type


Type Candidates
    CStudNum As String * 60
    CLastName As String * 60
    CFirstName As String * 60
    CMidName As String * 60
    CFullName As String * 200
    
    CCourse As String * 200
    CPos As String * 200
    CParty As String * 200
    CPic As String * 100
    
    CVote As Integer
    CAddVote As Integer
    
    
End Type


Type Courses
    Course As String * 200
End Type





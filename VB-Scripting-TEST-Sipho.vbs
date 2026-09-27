Option Explicit

Class Account

    public strAccountNumber
    public strAccountHolder
    public dblBalance

'============================================INTIALISE BALANCE SUB ROUTINE==========================================================
    Private Sub Class_Initialize()
            dblBalance = 1000.00
    End Sub
'============================================WITHDRAWAL FUNCTION==========================================================
    public Sub WithDrawal(dblAmount)
        'so we need to check is the amount is more than the balance then display an error message
        'we also need to check if the ablAmount is negative or if the balance is 0
        IF dblAmount <= 0 THEN
            Msgbox "You cannot withdraw zero or a negative amount"
        ELSEIF dblBalance = 0 THEN
            Msgbox "Insufficient Funds. Your Balance is 0 you cannot withdraw any funds"
        ELSEIF dblBalance < dblAmount THEN
            Msgbox "Insufficient Funds. Your Balance is less than your Withdrawal amount you cannot withdraw any funds"
        ELSE
            dblBalance = dblBalance - dblAmount
        END IF
    End Sub
'============================================DEPOSIT FUNCTION==========================================================
    public Sub Deposit(dblAmount)
        IF dblAmount <= 0 THEN
            Msgbox "You cannot deposit zero or a negative amount"
        ELSE
            dblBalance = dblBalance + dblAmount
        END IF
    End Sub

    public Sub Transfer(strAccountNumberToTransferTo, dblAmount)
       IF dblAmount <= 0 THEN
            Msgbox "You cannot transfer zero or a negative amount"
        ELSEIF dblBalance = 0 THEN
            Msgbox "Insufficient Funds. Your Balance is 0 you cannot transfer any funds"
        ELSEIF dblBalance < dblAmount THEN
            Msgbox "Insufficient Funds. Your Balance is less than your Transfer amount"
        ELSE
            dblBalance = dblBalance - dblAmount
        End If
    End Sub
'============================================GET ACCOUNT NUMBER FUNCTION==========================================================
    public Function getAccountNumber()
        getAccountNumber = "ACC-" & strAccountNumber
    End Function

'============================================GET ACCOUNT DETAILS FUNCTION==========================================================
    ' Takes no parameters and returns the account holder, account number, and balance nicely formatted
    public Function getAccountDetails()
        getAccountDetails = "Account Holder: " & strAccountHolder & vbNewLine & _
                                "Account Number: " & getAccountNumber & vbNewLine & _
                                "Remaining Balance: R" & dblBalance
    End Function

    '============================================CHECK BALANCE SUB ROUTINE==========================================================
    ' Displays the balance / account details on a message box
    public Sub CheckBalance()
        MsgBox getAccountDetails(), 0, "Account Balance & Details"
    End Sub





End Class

'===================================================MAIN PROGRAM=================================================
DIM objAcountForUser1
DIM dblwithdrawalamount
DIM dblDepositAmount
DIM straccHolder
DIM straccountNo
DIM strSelection
DIM strSelectionOutput
DIM strreceivingAcc, dbltransferAmt
SET objAcountForUser1 = NEW Account
'======================================================================================================================================
'start of by asking the user to capture their account number
DO WHILE TRUE
    straccountNo = InputBox("Please enter your Account Number (7-10 digits eg 1234567): ", "Account Login")
  
    IF straccountNo = "" THEN
        Msgbox "You Need to Enter an Account Number to proceed",0,"ENTER AN ACCOUNT NUMBER"
    ELSEIF IsNumeric(straccountNo) And Len(straccountNo) >= 7 And Len(straccountNo) <= 10 THEN
        objAcountForUser1.strAccountNumber = straccountNo
        EXIT DO
    ELSE
        Msgbox "Your Account Number Must Only Be Numbers and be between 7 and 10 digits long.",0,"INVALID ACCOUNT NUMBER"
    END IF
LOOP
'======================================================================================================================================
'start of by asking the user to capture their full name and surname to intialise account holder
DO WHILE TRUE
    straccHolder = InputBox("Please enter your Full Name and Surname(eg Sipho Sehlapelo): ", "Account Login")
  
    IF straccHolder = "" THEN
        Msgbox "You Need to Enter your Full Name to proceed",0,"ENTER YOUR FULL NAME"
    ELSEIF NOT IsNumeric(straccHolder) THEN
        objAcountForUser1.strAccountHolder = straccHolder
        EXIT DO
    ELSE
        Msgbox "Enter a Valid Full Name",0,"INVALID NAME"
    END IF
LOOP
'===================================================================================================================================
'then ask them to select one of those options
strSelectionOutput = "Select a number accordding to what you want: " & vbNewLine & _
                    "1 - Withdrawal" & vbNewLine & _
                    "2 - Deposit" & vbNewLine & _
                    "3 - Transfer" & vbNewLine & _
                    "4 - Check balance" & vbNewLine & _
                    "5 - Exit "

    
DO WHILE TRUE
    strSelection = InputBox(strSelectionOutput, "Main Menu")
    IF strSelection = "" THEN 
        EXIT DO ' Exit program if Cancel is clicked
    ELSEIF NOT IsNumeric(strSelection) THEN
        MsgBox "Please Enter a valid number"
    ELSE
        SELECT CASE CInt(strSelection)
        case 1
            'ask the user to enter the withdrawal amount 
            DO WHILE TRUE
                dblwithdrawalamount = InputBox("Please Enter a WithDrawal Amount", "WithDrawal Menu")
                IF dblwithdrawalamount = "" THEN
                    EXIT DO
                ELSEIF NOT IsNumeric(dblwithdrawalamount) THEN
                   MsgBox "Please Enter a valid number", 0, "Invalid Input"
                ELSE
                    objAcountForUser1.WithDrawal CDbl(dblwithdrawalamount)
                    objAcountForUser1.CheckBalance
                    EXIT DO
                END IF
            LOOP
        case 2
            'ask user to enter the deposit amount and make the neccessary changes to the object
            DO WHILE TRUE
                dblDepositAmount = InputBox("Please Enter a Deposit Amount", "Deposit Menu")
                IF dblDepositAmount = "" THEN
                    EXIT DO
                ELSEIF NOT IsNumeric(dblDepositAmount) THEN
                    MsgBox "Please Enter a valid number"
                ELSE
                    objAcountForUser1.Deposit CDbl(dblDepositAmount)
                    objAcountForUser1.CheckBalance
                    Exit Do
                END IF
            LOOP
        case 3
           DO WHILE TRUE
                strreceivingAcc = InputBox("Enter the receiving account number (7-10 digits):", "Transfer Menu")
                
                IF strreceivingAcc = "" THEN 
                    EXIT DO
                ELSEIF NOT IsNumeric(strreceivingAcc) OR Len(strreceivingAcc) < 7 OR Len(strreceivingAcc) > 10 THEN
                    Msgbox "Account Number Must Only Be Numbers and be between 7 and 10 digits long.", 0, "INVALID ACCOUNT NUMBER"
                ELSE
                    ' Capture into your old variable name as a string first
                    dbltransferAmt = InputBox("Enter amount to transfer:", "Transfer Menu")
                    
                    IF dbltransferAmt = "" THEN
                        EXIT DO
                    ELSEIF NOT IsNumeric(dbltransferAmt) THEN
                        MsgBox "Please enter a valid numeric amount.", 0, "Invalid Input"
                    ELSE
                        objAcountForUser1.Transfer strreceivingAcc, CDbl(dbltransferAmt)
                        objAcountForUser1.CheckBalance
                        EXIT DO
                    END IF
                END IF
            LOOP
        case 4
            objAcountForUser1.CheckBalance
        case 5
            EXIT DO
        case else
            MsgBox "Please Enter a valid number from the options"
        END SELECT
    END IF
LOOP




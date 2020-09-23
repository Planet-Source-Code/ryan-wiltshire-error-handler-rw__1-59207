VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum ErrorAction
    DoNothing = 0
    ExitProcedure = 1
    ExitProgram = 2
    ExitForm = 3
    ResumeAction = 4
    ResumeNext = 5
End Enum
Public Function ErrorHandler(Optional ByVal sProcedure As String) As ErrorAction
Dim sMessage As String
Dim sTitle As String

    If IsMissing(sProcedure) Or sProcedure = vbNullString Then sProcedure = Err.Source
    sTitle = IIf(Err.Number > 30000, "Application Error!", "Visual Basic Error!")
      
    sMessage = "The following " & IIf(Err.Number > 30000, "Application ", "Visual Basic ") & "error has occurred:" & vbCrLf & vbCrLf

    sMessage = sMessage & "Error Number:" & vbTab & Err.Number & vbCrLf
    sMessage = sMessage & "Description:" & vbTab & Err.Description & vbCrLf
    sMessage = sMessage & "Source:" & vbTab & vbTab & sProcedure & vbCrLf
    
    sMessage = sMessage & vbCrLf
    sMessage = sMessage & "Contact your system administrator or your local support desk, and quote the message above." & vbCrLf
      
    MsgBox sMessage, vbInformation
    ErrorHandler = GetAction
End Function
Private Function GetAction() As ErrorAction
    Select Case Err.Number
        Case 3              'Return without GoSub
            GetAction = ExitProcedure
        Case 5              'Invalid procedure call
            GetAction = ExitProcedure
        Case 6              'Overflow
            GetAction = ExitForm
        Case 7              'Out of memory
            GetAction = ExitForm
        Case 9              'Subscript out of range
            GetAction = ExitProcedure
        Case 10             'This array is fixed or temporarily locked
            GetAction = ExitProcedure
        Case 11             'Division by zero
            GetAction = ExitProcedure
        Case 13             'Type mismatch
            GetAction = ExitProcedure
        Case 14             'Out of string space
            GetAction = ExitForm
        Case 16             'Expression too complex
            GetAction = ExitForm
        Case 17             'Can't perform requested operation
            GetAction = ExitForm
        Case 18             'User interrupt occurred
            GetAction = ExitProcedure
        Case 20             'Resume without error
            GetAction = ExitForm
        Case 28             'Out of stack space
            GetAction = ExitForm
        Case 35             'Sub, Function, or Property not defined
            GetAction = ExitForm
        Case 47             'Too many DLL application clients
            GetAction = ExitForm
        Case 48             'Error in loading DLL
            GetAction = ExitForm
        Case 49             'Bad DLL calling convention
            GetAction = ExitForm
        Case 51             'Internal error
            GetAction = ExitProcedure
        Case 52             'Bad file name or number
            GetAction = ExitProcedure
        Case 53             'File not found
            GetAction = ExitProcedure
        Case 54             'Bad file mode
            GetAction = ExitProcedure
        Case 55             'File already open
            GetAction = ExitProcedure
        Case 57             'Device I/O error
            GetAction = ExitProcedure
        Case 58             'File already exists
            GetAction = ExitProcedure
        Case 59             'Bad record length
            GetAction = ExitProcedure
        Case 61             'Disk full
            GetAction = ExitProcedure
        Case 62             'Input past end of file
            GetAction = ExitProcedure
        Case 63             'Bad record number
            GetAction = ExitProcedure
        Case 67             'Too many files
            GetAction = ExitProcedure
        Case 68             'Device unavailable
            GetAction = ExitProcedure
        Case 70             'Permission denied
            GetAction = ExitProcedure
        Case 71             'Disk not ready
            GetAction = ExitProcedure
        Case 74             'Can't rename with different drive
            GetAction = ExitProcedure
        Case 75             'Path/File access error
            GetAction = ExitProcedure
        Case 76             'Path not found
            GetAction = ExitProcedure
        Case 91             'Object variable or With block variable not set
            GetAction = ExitProcedure
        Case 92             'For loop not initialized
            GetAction = ExitProcedure
        Case 93             'Invalid pattern string
            GetAction = ExitForm
        Case 94             'Invalid use of Null
            GetAction = ResumeNext
        Case 97             'Can't call Friend procedure on an object that is not an instance of the defining class
            GetAction = ExitForm
        Case 98             'A property or method call cannot include a reference to a private object, either as an argument or as a return value
            GetAction = ExitForm
        Case 298            'System DLL could not be loaded
            GetAction = ExitProgram
        Case 320            'Can't use character device names in specified file names
            GetAction = ExitForm
        Case 321            'Invalid file format
            GetAction = ExitForm
        Case 322            'Can’t create necessary temporary file
            GetAction = ExitForm
        Case 325            'Invalid format in resource file
            GetAction = ExitForm
        Case 327            'Data value named not found
            GetAction = ExitForm
        Case 328            'Illegal parameter; can't write arrays
            GetAction = ExitForm
        Case 335            'Could not access system registry
            GetAction = ExitForm
        Case 336            'Component not correctly registered
            GetAction = ExitForm
        Case 337            'Component not found
            GetAction = ExitForm
        Case 338            'Component did not run correctly
            GetAction = ExitForm
        Case 360            'Object already loaded
            GetAction = ResumeNext
        Case 361            'Can't load or unload this object
            GetAction = ExitForm
        Case 363            'Control specified not found
            GetAction = ExitForm
        Case 364            'Object was unloaded
            GetAction = ExitProcedure
        Case 365            'Unable to unload within this context
            GetAction = ExitForm
        Case 368            'The specified file is out of date. This program requires a later version
            GetAction = ExitForm
        Case 371            'The specified object can't be used as an owner form for Show
            GetAction = ExitForm
        Case 380            'Invalid property value
            GetAction = ExitForm
        Case 381            'Invalid property-array index
            GetAction = ExitForm
        Case 382            'Property Set can't be executed at run time
            GetAction = ExitForm
        Case 383            'Property Set can't be used with a read-only property
            GetAction = ExitForm
        Case 385            'Need property-array index
            GetAction = ExitForm
        Case 387            'Property Set not permitted
            GetAction = ExitForm
        Case 393            'Property Get can't be executed at run time
            GetAction = ExitForm
        Case 394            'Property Get can't be executed on write-only property
            GetAction = ExitForm
        Case 400            'Form already displayed; can't show modally
            GetAction = ResumeNext
        Case 402            'Code must close topmost modal form first
            GetAction = ResumeNext
        Case 419            'Permission to use object denied
            GetAction = ExitForm
        Case 422            'Property not found
            GetAction = ExitForm
        Case 423            'Property or method not found
            GetAction = ExitForm
        Case 424            'Object required
            GetAction = ExitProcedure
        Case 425            'Invalid object use
            GetAction = ExitForm
        Case 429            'Component can't create object or return reference to this object
            GetAction = ExitForm
        Case 430            'Class doesn't support Automation
            GetAction = ExitProgram
        Case 432            'File name or class name not found during Automation operation
            GetAction = ExitProgram
        Case 438            'Object doesn't support this property or method
            GetAction = ExitForm
        Case 440            'Automation error
            GetAction = ExitForm
        Case 442            'Connection to type library or object library for remote process has been lost
            GetAction = ExitForm
        Case 443            'Automation object doesn't have a default value
            GetAction = ExitForm
        Case 445            'Object doesn't support this action
            GetAction = ExitForm
        Case 446            'Object doesn't support named arguments
            GetAction = ExitForm
        Case 447            'Object doesn't support current locale setting
            GetAction = ExitForm
        Case 448            'Named argument not found
            GetAction = ExitProcedure
        Case 449            'Argument not optional or invalid property assignment
            GetAction = ExitProcedure
        Case 450            'Wrong number of arguments or invalid property assignment
            GetAction = ExitProcedure
        Case 451            'Object not a collection
            GetAction = ExitProcedure
        Case 452            'Invalid ordinal
            GetAction = ExitForm
        Case 453            'Specified not found
            GetAction = ExitForm
        Case 454            'Code resource not found
            GetAction = ExitForm
        Case 455            'Code resource lock error
            GetAction = ExitForm
        Case 457            'This key is already associated with an element of this collection
            GetAction = ExitForm
        Case 458            'Variable uses a type not supported in Visual Basic
            GetAction = ExitProcedure
        Case 459            'This component doesn't support the set of events
            GetAction = ExitForm
        Case 460            'Invalid Clipboard format
            GetAction = ExitForm
        Case 461            'Method or data member not found
            GetAction = ExitForm
        Case 462            'The remote server machine does not exist or is unavailable
            GetAction = ExitProgram
        Case 463            'Class not registered on local machine
            GetAction = ExitProgram
        Case 480            'Can't create AutoRedraw image
            GetAction = ExitForm
        Case 481            'Invalid picture
            GetAction = ExitProcedure
        Case 482            'Printer error
            GetAction = ExitProcedure
        Case 483            'Printer driver does not support specified property
            GetAction = ExitProcedure
        Case 484            'Problem getting printer information from the system. Make sure the printer is set up correctly
            GetAction = ExitProcedure
        Case 485            'Invalid picture type
            GetAction = ExitProcedure
        Case 486            'Can't print form image to this type of printer
            GetAction = ExitProcedure
        Case 520            'Can't empty Clipboard
            GetAction = ExitProcedure
        Case 521            'Can't open Clipboard
            GetAction = ExitProcedure
        Case 735            'Can't save file to TEMP directory
            GetAction = ExitProcedure
        Case 744            'Search text not found
            GetAction = ResumeNext
        Case 746            'Replacements too long
            GetAction = ExitForm
        Case 2424     'Unknown name
            GetAction = ExitProcedure
        Case 2425     'Unknown function name
            GetAction = ExitProcedure
        Case 2426     'Function isn't available in expressions
            GetAction = ExitProcedure
        Case 2427     'Object has no value
            GetAction = ExitProcedure
        Case 2428     'Invalid arguments used with domain function
            GetAction = ExitProcedure
        Case 2429     'In operator without ()
            GetAction = ExitProcedure
        Case 2430     'Between operator without And
            GetAction = ExitProcedure
        Case 2431     'Syntax error
            GetAction = ExitProcedure
        Case 2432     'Syntax error
            GetAction = ExitProcedure
        Case 2433     'Syntax error
            GetAction = ExitProcedure
        Case 2434     'Syntax error
            GetAction = ExitProcedure
        Case 2435     'Extra )
            GetAction = ExitProcedure
        Case 2436     'Missing ), ], or
            GetAction = ExitProcedure
        Case 2437     'Invalid use of vertical bars
            GetAction = ExitProcedure
        Case 2438     'Syntax error
            GetAction = ExitProcedure
        Case 2439     'Wrong number of arguments used with function
            GetAction = ExitProcedure
        Case 2440     'IIF function without ()
            GetAction = ExitProcedure
        Case 2442     'Invalid use of parentheses
            GetAction = ExitProcedure
        Case 2443     'Invalid use of Is operator
            GetAction = ExitProcedure
        Case 2445     'Expression too complex
            GetAction = ExitProcedure
        Case 2446     'Out of memory during calculation
            GetAction = ExitProcedure
        Case 2447     'Invalid use of '.', '!', or '()'.
            GetAction = ExitProcedure
        Case 2448     'Can't set value.
            GetAction = ExitProcedure
        Case 2449     'Invalid method in expression.
            GetAction = ExitProcedure
        Case 2450     'Invalid reference to form ' '.
            GetAction = ExitProcedure
        Case 2451     'Invalid reference to report ' '.
            GetAction = ExitProcedure
        Case 2452     'Invalid reference to Parent property.
            GetAction = ExitProcedure
        Case 2453     'Invalid reference to control ' '.
            GetAction = ExitProcedure
        Case 2454     'Invalid reference to '! '.
            GetAction = ExitProcedure
        Case 2455     'Invalid reference to property ' '.
            GetAction = ExitProcedure
        Case 2456     'Invalid form number reference.
            GetAction = ExitProcedure
        Case 2457     'Invalid report number reference.
            GetAction = ExitProcedure
        Case 2458     'Invalid control number reference.
            GetAction = ExitProcedure
        Case 2459     'Can't refer to Parent property in Design view.
            GetAction = ExitProcedure
        Case 2460     'Can't refer to Dynaset property in Design view.
            GetAction = ExitProcedure
        Case 2461     'Invalid section reference.
            GetAction = ExitProcedure
        Case 2462     'Invalid section number reference.
            GetAction = ExitProcedure
        Case 2463     'Invalid group level reference.
            GetAction = ExitProcedure
        Case 2464     'Invalid group level number reference.
            GetAction = ExitProcedure
        Case 2465     'Invalid reference to field ' '.
            GetAction = ExitProcedure
        Case 2466     'Invalid reference to Dynaset property.
            GetAction = ExitProcedure
        Case 2467     'Object referred to in expression no longer exists.
            GetAction = ExitProcedure
        Case 2468     'Invalid argument used with DatePart, DateAdd or DateDiff function.
            GetAction = ExitProcedure
        Case 2469     '1 in validation rule: '|2'.
            GetAction = ExitProcedure
        Case 2470     'in validation rule.
            GetAction = ExitProcedure
        Case 2471     'in query.
            GetAction = ExitProcedure
        Case 2472     'in linked master field.
            GetAction = ExitProcedure
        Case 2473     '1 in '|2' expression.
            GetAction = ExitProcedure
        Case 2474     'No control is active.
            GetAction = ExitProcedure
        Case 2475     'No form is active.
            GetAction = ExitProcedure
        Case 2476     'No report is active.
            GetAction = ExitProcedure
        Case 2477     'Invalid subclass ' ' referred to in TypeOf function.
            GetAction = ExitProcedure
        Case 3000     'Reserved error ( ); there is no message for this error.
            GetAction = ExitProcedure
        Case 3001     'Invalid argument.
            GetAction = ExitProcedure
        Case 3002     'Couldn't start session.
            GetAction = ExitProcedure
        Case 3003     'Couldn't start transaction; too many transactions already nested.
            GetAction = ExitProcedure
        Case 3004     'Couldn't find database ' '.
            GetAction = ExitProcedure
        Case 3005     '' ' isn't a valid database name.
            GetAction = ExitProcedure
        Case 3006     'Database ' ' is exclusively locked.
            GetAction = ExitProcedure
        Case 3007     'Couldn't open database ' '.
            GetAction = ExitProcedure
        Case 3008     'Table ' ' is exclusively locked.
            GetAction = ExitProcedure
        Case 3009     'Couldn't lock table ' '; currently in use.
            GetAction = ExitProcedure
        Case 3010     'Table ' ' already exists.
            GetAction = ExitProcedure
        Case 3011     'Couldn't find object ' '.
            GetAction = ExitProcedure
        Case 3012     'Object ' ' already exists.
            GetAction = ExitProcedure
        Case 3013     'Couldn't rename installable ISAM file.
            GetAction = ExitProcedure
        Case 3014     'Can't open any more tables.
            GetAction = ExitProcedure
        Case 3015     '' ' isn't an index in this table.
            GetAction = ExitProcedure
        Case 3016     'Field won't fit in record.
            GetAction = ExitProcedure
        Case 3017     'Field length is too long.
            GetAction = ExitProcedure
        Case 3018     'Couldn't find field ' '.
            GetAction = ExitProcedure
        Case 3019     'Operation invalid without a current index.
            GetAction = ExitProcedure
        Case 3020     'Update without AddNew or Edit.
            GetAction = ExitProcedure
        Case 3021           'EOF OR BOF
            GetAction = ExitProcedure
        Case 3022     'Can't have duplicate key; index changes were unsuccessful.
            GetAction = ExitProcedure
        Case 3023     'AddNew or Edit already used.
            GetAction = ExitProcedure
        Case 3024     'Couldn't find file ' '.
            GetAction = ExitProcedure
        Case 3025     'Can't open any more files.
            GetAction = ExitProcedure
        Case 3026     'Not enough space on disk.
            GetAction = ExitProcedure
        Case 3027     'Couldn't update; database is read-only.
            GetAction = ExitProcedure
        Case 3028     'Couldn't initialize data access because file 'SYSTEM.MDA' couldn't be opened.
            GetAction = ExitProcedure
        Case 3029     'Not a valid account name or password.
            GetAction = ExitProcedure
        Case 3030     '' ' isn't a valid account name.
            GetAction = ExitProcedure
        Case 3031     'Not a valid password.
            GetAction = ExitProcedure
        Case 3032     'Can't delete account.
            GetAction = ExitProcedure
        Case 3033     'No permission for ' '.
            GetAction = ExitProcedure
        Case 3034     'Commit or Rollback without BeginTrans.
            GetAction = ExitProcedure
        Case 3035     'Out of memory.
            GetAction = ExitProcedure
        Case 3036     'Database has reached maximum size.
            GetAction = ExitProcedure
        Case 3037     'Can't open any more tables or queries.
            GetAction = ExitProcedure
        Case 3038     'Out of memory.
            GetAction = ExitProcedure
        Case 3039     'Couldn't create index; too many indexes already defined.
            GetAction = ExitProcedure
        Case 3040     'Disk I/O error during read.
            GetAction = ExitProcedure
        Case 3041     'Incompatible database version.
            GetAction = ExitProcedure
        Case 3042     'Out of MS-DOS file handles.
            GetAction = ExitProcedure
        Case 3043     'Disk or network error.
            GetAction = ExitProcedure
        Case 3044     '' ' isn't a valid path.
            GetAction = ExitProcedure
        Case 3045     'Couldn't use ' '; file already in use.
            GetAction = ExitProcedure
        Case 3046     'Couldn't save; currently locked by another user.
            GetAction = ExitProcedure
        Case 3047     'Record is too large.
            GetAction = ExitProcedure
        Case 3048     'Can't open any more databases.
            GetAction = ExitProcedure
        Case 3049     '' ' is corrupted or isn't a Microsoft Access database.
            GetAction = ExitProcedure
        Case 3050     'Couldn't lock file; SHARE.EXE hasn't been loaded.
            GetAction = ExitProcedure
        Case 3051     'Couldn't open file ' '.
            GetAction = ExitProcedure
        Case 3052     'MS-DOS file sharing lock count exceeded.  You need to increase the number of locks installed with SHARE.EXE.
            GetAction = ExitProcedure
        Case 3053     'Too many client tasks.
            GetAction = ExitProcedure
        Case 3054     'Too many Memo or Long Binary fields.
            GetAction = ExitProcedure
        Case 3055     'Not a valid file name.
            GetAction = ExitProcedure
        Case 3056     'Couldn't repair this database.
            GetAction = ExitProcedure
        Case 3057     'Operation not supported on attached tables.
            GetAction = ExitProcedure
        Case 3058     'Can't have Null value in index.
            GetAction = ExitProcedure
        Case 3059     'Operation canceled by user.
            GetAction = ExitProcedure
        Case 3060     'Wrong data type for parameter ' '.
            GetAction = ExitProcedure
        Case 3061     '1 parameters were expected, but only |2 were supplied.
            GetAction = ExitProcedure
        Case 3062     'Duplicate output alias ' '.
            GetAction = ExitProcedure
        Case 3063     'Duplicate output destination ' '.
            GetAction = ExitProcedure
        Case 3064     'Can't open action query ' '.
            GetAction = ExitProcedure
        Case 3065     'Can't execute a non-action query.
            GetAction = ExitProcedure
        Case 3066     'Query must have at least one output field.
            GetAction = ExitProcedure
        Case 3067     'Query input must contain at least one table or query.
            GetAction = ExitProcedure
        Case 3068     'Not a valid alias name.
            GetAction = ExitProcedure
        Case 3069     'Can't have action query ' ' as an input.
            GetAction = ExitProcedure
        Case 3070     'Can't bind name ' '.
            GetAction = ExitProcedure
        Case 3071     'Can't evaluate expression.
            GetAction = ExitProcedure
        Case 3073     'Operation must use an updatable query.
            GetAction = ExitProcedure
        Case 3074     'Can't repeat table name ' ' in FROM clause.
            GetAction = ExitProcedure
        Case 3075     '1 in query expression '|2'.
            GetAction = ExitProcedure
        Case 3076     'in criteria expression.
            GetAction = ExitProcedure
        Case 3077     'in expression.
            GetAction = ExitProcedure
        Case 3078     'Couldn't find input table or query ' '.
            GetAction = ExitProcedure
        Case 3079     'Ambiguous field reference ' '.
            GetAction = ExitProcedure
        Case 3080     'Joined table ' ' not listed in FROM clause.
            GetAction = ExitProcedure
        Case 3081     'Can't join more than one table with the same name ( ).
            GetAction = ExitProcedure
        Case 3082     'JOIN operation ' ' refers to a non-joined table.
            GetAction = ExitProcedure
        Case 3083     'Can't use internal report query.
            GetAction = ExitProcedure
        Case 3084     'Can't insert into action query.
            GetAction = ExitProcedure
        Case 3085     'Undefined function ' ' in expression.
            GetAction = ExitProcedure
        Case 3086     'Couldn't delete from specified tables.
            GetAction = ExitProcedure
        Case 3087     'Too many expressions in GROUP BY clause.
            GetAction = ExitProcedure
        Case 3088     'Too many expressions in ORDER BY clause.
            GetAction = ExitProcedure
        Case 3089     'Too many expressions in DISTINCT output.
            GetAction = ExitProcedure
        Case 3090     'Resultant table may not have more than one Counter field.
            GetAction = ExitProcedure
        Case 3091     'HAVING clause ( ) without grouping or aggregation.
            GetAction = ExitProcedure
        Case 3092     'Can't use HAVING clause in TRANSFORM statement.
            GetAction = ExitProcedure
        Case 3093     'ORDER BY clause ( ) conflicts with DISTINCT.
            GetAction = ExitProcedure
        Case 3094     'ORDER BY clause ( ) conflicts with GROUP BY clause.
            GetAction = ExitProcedure
        Case 3095     'Can't have aggregate function in expression ( ).
            GetAction = ExitProcedure
        Case 3096     'Can't have aggregate function in WHERE clause ( ).
            GetAction = ExitProcedure
        Case 3097     'Can't have aggregate function in ORDER BY clause ( ).
            GetAction = ExitProcedure
        Case 3098     'Can't have aggregate function in GROUP BY clause ( ).
            GetAction = ExitProcedure
        Case 3099     'Can't have aggregate function in JOIN operation ( ).
            GetAction = ExitProcedure
        Case 3100     'Can't set field ' ' in join key to Null.
            GetAction = ExitProcedure
        Case 3101     'Join is broken by value(s) in fields ' '.
            GetAction = ExitProcedure
        Case 3102     'Circular reference caused by ' '.
            GetAction = ExitProcedure
        Case 3103     'Circular reference caused by alias ' ' in query definition's SELECT list.
            GetAction = ExitProcedure
        Case 3104     'Can't specify Fixed Column Heading ' ' in a crosstab query more than once.
            GetAction = ExitProcedure
        Case 3105     'Missing destination field name in SELECT INTO statement ( ).
            GetAction = ExitProcedure
        Case 3106     'Missing destination field name in UPDATE statement ( ).
            GetAction = ExitProcedure
        Case 3107     'Couldn't insert; no insert permission for table or query ' '.
            GetAction = ExitProcedure
        Case 3108     'Couldn't replace; no replace permission for table or query ' '.
            GetAction = ExitProcedure
        Case 3109     'Couldn't delete; no delete permission for table or query ' '.
            GetAction = ExitProcedure
        Case 3110     'Couldn't read definitions; no read definitions permission for table or query ' '.
            GetAction = ExitProcedure
        Case 3111     'Couldn't create; no create permission for table or query ' '.
            GetAction = ExitProcedure
        Case 3112     'Couldn't read; no read permission for table or query ' '.
            GetAction = ExitProcedure
        Case 3113     'Can't update ' '; field not updatable.
            GetAction = ExitProcedure
        Case 3114     'Can't include Memo or Long Binary when you select unique values( ).
            GetAction = ExitProcedure
        Case 3115     'Can't have Memo or Long Binary in aggregate argument ( ).
            GetAction = ExitProcedure
        Case 3116     'Can't have Memo or Long Binary in criteria ( ) for aggregate function.
            GetAction = ExitProcedure
        Case 3117     'Can't sort on Memo or Long Binary ( ).
            GetAction = ExitProcedure
        Case 3118     'Can't join on Memo or Long Binary ( ).
            GetAction = ExitProcedure
        Case 3119     'Can't group on Memo or Long Binary ( ).
            GetAction = ExitProcedure
        Case 3120     'Can't group on fields selected with '*' ( ).
            GetAction = ExitProcedure
        Case 3121     'Can't group on fields selected with '*'.
            GetAction = ExitProcedure
        Case 3122     '' ' not part of aggregate function or grouping.
            GetAction = ExitProcedure
        Case 3123     'Can't use '*' in crosstab query.
            GetAction = ExitProcedure
        Case 3124     'Can't input from internal report query ( ).
            GetAction = ExitProcedure
        Case 3125     '' ' isn't a valid name.
            GetAction = ExitProcedure
        Case 3126     'Invalid bracketing of name ' '.
            GetAction = ExitProcedure
        Case 3127     'INSERT INTO statement contains unknown field name ' '.
            GetAction = ExitProcedure
        Case 3128     'Must specify tables to delete from.
            GetAction = ExitProcedure
        Case 3129     'Invalid SQL statement; expected 'DELETE', 'INSERT', 'PROCEDURE', 'SELECT', or 'UPDATE'.
            GetAction = ExitProcedure
        Case 3130     'Syntax error in DELETE statement.
            GetAction = ExitProcedure
        Case 3131     'Syntax error in FROM clause.
            GetAction = ExitProcedure
        Case 3132     'Syntax error in GROUP BY clause.
            GetAction = ExitProcedure
        Case 3133     'Syntax error in HAVING clause.
            GetAction = ExitProcedure
        Case 3134     'Syntax error in INSERT statement.
            GetAction = ExitProcedure
        Case 3135     'Syntax error in JOIN operation.
            GetAction = ExitProcedure
        Case 3136     'Syntax error in LEVEL clause.
            GetAction = ExitProcedure
        Case 3137     'Missing semicolon (;) at end of SQL statement.
            GetAction = ExitProcedure
        Case 3138     'Syntax error in ORDER BY clause.
            GetAction = ExitProcedure
        Case 3139     'Syntax error in PARAMETER clause.
            GetAction = ExitProcedure
        Case 3140     'Syntax error in PROCEDURE clause.
            GetAction = ExitProcedure
        Case 3141     'Syntax error in SELECT statement.
            GetAction = ExitProcedure
        Case 3142     'Characters found after end of SQL statement.
            GetAction = ExitProcedure
        Case 3143     'Syntax error in TRANSFORM statement.
            GetAction = ExitProcedure
        Case 3144     'Syntax error in UPDATE statement.
            GetAction = ExitProcedure
        Case 3145     'Syntax error in WHERE clause.
            GetAction = ExitProcedure
        Case 3146     'ODBC--call failed.
            GetAction = ExitProcedure
        Case 3147     'ODBC--data buffer overflow.
            GetAction = ExitProcedure
        Case 3148     'ODBC--connection failed.
            GetAction = ExitProcedure
        Case 3149     'ODBC--incorrect DLL.
            GetAction = ExitProcedure
        Case 3150     'ODBC--missing DLL.
            GetAction = ExitProcedure
        Case 3151     'ODBC--connection to ' ' failed.
            GetAction = ExitProcedure
        Case 3152     'ODBC--incorrect driver version ' 1'; expected version '|2'.
            GetAction = ExitProcedure
        Case 3153     'ODBC--incorrect server version ' 1'; expected version '|2'.
            GetAction = ExitProcedure
        Case 3154     'ODBC--couldn't find DLL ' '.
            GetAction = ExitProcedure
        Case 3155     'ODBC--insert failed.
            GetAction = ExitProcedure
        Case 3156     'ODBC--delete failed.
            GetAction = ExitProcedure
        Case 3157     'ODBC--update failed.
            GetAction = ExitProcedure
        Case 3158     'Couldn't save record; currently locked by another user.
            GetAction = ExitProcedure
        Case 3159     'Not a valid bookmark.
            GetAction = ExitProcedure
        Case 3160     'Table isn't open.
            GetAction = ExitProcedure
        Case 3161     'Couldn't decrypt file.
            GetAction = ExitProcedure
        Case 3162     'Null is invalid.
            GetAction = ExitProcedure
        Case 3163     'Couldn't insert or paste; data too long for field.
            GetAction = ExitProcedure
        Case 3164     'Couldn't update field.
            GetAction = ExitProcedure
        Case 3165     'Couldn't open .INF file.
            GetAction = ExitProcedure
        Case 3166     'Missing memo file.
            GetAction = ExitProcedure
        Case 3167     'Record is deleted.
            GetAction = ExitProcedure
        Case 3168     'Invalid .INF file.
            GetAction = ExitProcedure
        Case 3169     'Illegal type in expression.
            GetAction = ExitProcedure
        Case 3170     'Couldn't find installable ISAM.
            GetAction = ExitProcedure
        Case 3171     'Couldn't find net path or user name.
            GetAction = ExitProcedure
        Case 3172     'Couldn't open PARADOX.NET.
            GetAction = ExitProcedure
        Case 3173     'Couldn't open table 'MSysAccounts' in SYSTEM.MDA.
            GetAction = ExitProcedure
        Case 3174     'Couldn't open table 'MSysGroups' in SYSTEM.MDA.
            GetAction = ExitProcedure
        Case 3175     'Date is out of range or is in an invalid format.
            GetAction = ExitProcedure
        Case 3176     'Couldn't open file ' '.
            GetAction = ExitProcedure
        Case 3177     'Not a valid table name.
            GetAction = ExitProcedure
        Case 3178     'Out of memory.
            GetAction = ExitProcedure
        Case 3179     'Encountered unexpected end of file.
            GetAction = ExitProcedure
        Case 3180     'Couldn't write to file ' '.
            GetAction = ExitProcedure
        Case 3181     'Invalid range.
            GetAction = ExitProcedure
        Case 3182     'Invalid file format.
            GetAction = ExitProcedure
        Case 3183     'Not enough space on temporary disk.
            GetAction = ExitProcedure
        Case 3184     'Couldn't execute query; couldn't find linked table.
            GetAction = ExitProcedure
        Case 3185     'SELECT INTO remote database tried to produce too many fields.
            GetAction = ExitProcedure
        Case 3186     'Couldn't save; currently locked by user ' 2' on machine '|1'.
            GetAction = ExitProcedure
        Case 3187     'Couldn't read; currently locked by user ' 2' on machine '|1'.
            GetAction = ExitProcedure
        Case 3188     'Couldn't update; currently locked by another session on this machine.
            GetAction = ExitProcedure
        Case 3189     'Table ' 1' is exclusively locked by user '|3' on machine '|2'.
            GetAction = ExitProcedure
        Case 3190     'Too many fields defined.
            GetAction = ExitProcedure
        Case 3191     'Can't define field more than once.
            GetAction = ExitProcedure
        Case 3192     'Couldn't find output table ' '.
            GetAction = ExitProcedure
        Case 3193     '(unknown)
            GetAction = ExitProcedure
        Case 3194     '(unknown)
            GetAction = ExitProcedure
        Case 3195     '(expression)
            GetAction = ExitProcedure
        Case 3196     'Couldn't use ' '; database already in use.
            GetAction = ExitProcedure
        Case 3197     'Data has changed; operation stopped.
            GetAction = ExitProcedure
        Case 3198     'Couldn't start session.  Too many sessions already active.
            GetAction = ExitProcedure
        Case 3199     'Couldn't find reference.
            GetAction = ExitProcedure
        Case 3200     'Can't delete or change record. Since related records exist in table, referential integrity rules would be violated.
            GetAction = ExitProcedure
        Case 3201     'Can't add or change record.  Referential integrity rules require related record in table ' '.
            GetAction = ExitProcedure
        Case 3202     'Couldn't save; currently locked by another user.
            GetAction = ExitProcedure
        Case 3203     'Can't specify subquery in expression ( ).
            GetAction = ExitProcedure
        Case 3204     'Database already exists.
            GetAction = ExitProcedure
        Case 3205     'Too many crosstab column headers ( ).
            GetAction = ExitProcedure
        Case 3206     'Can't create a relationship between a field and itself.
            GetAction = ExitProcedure
        Case 3207     'Operation not supported on Paradox table with no primary key.
            GetAction = ExitProcedure
        Case 3208     'Invalid Deleted entry in [dBASE ISAM] section in INI file.
            GetAction = ExitProcedure
        Case 3209     'Invalid Stats entry in [dBASE ISAM] section in INI file.
            GetAction = ExitProcedure
        Case 3210     'Connect string too long.
            GetAction = ExitProcedure
        Case 3211     'Couldn't lock table ' '; currently in use.
            GetAction = ExitProcedure
        Case 3212     'Couldn't lock table ' 1'; currently in use by user '|3' on machine '|2'.
            GetAction = ExitProcedure
        Case 3213     'Invalid Date entry in [dBASE ISAM] section in INI file.
            GetAction = ExitProcedure
        Case 3214     'Invalid Mark entry in [dBASE ISAM] section in INI file.
            GetAction = ExitProcedure
        Case 3215     'Too many Btrieve tasks.
            GetAction = ExitProcedure
        Case 3216     'Parameter ' ' specified where a table name is required.
            GetAction = ExitProcedure
        Case 3217     'Parameter ' ' specified where a database name is required.
            GetAction = ExitProcedure
        Case 3218     'Couldn't update; currently locked.
            GetAction = ExitProcedure
        Case 3219     'Can't perform operation; it is illegal.
            GetAction = ExitProcedure
        Case 3220     'Wrong Paradox sort sequence.
            GetAction = ExitProcedure
        Case 3221     'Invalid entries in [Btrieve ISAM] section in WIN.INI.
            GetAction = ExitProcedure
        Case 3222     'Query can't contain a Database parameter.
            GetAction = ExitProcedure
        Case 3223     '' ' isn't a valid parameter name.
            GetAction = ExitProcedure
        Case 3224     'Btrieve--data dictionary is corrupted.
            GetAction = ExitProcedure
        Case 3225     'Encountered record locking deadlock while performing Btrieve operation.
            GetAction = ExitProcedure
        Case 3226     'Errors encountered while using the Btrieve DLL.
            GetAction = ExitProcedure
        Case 3227     'Invalid Century entry in [dBASE ISAM] section in INI file.
            GetAction = ExitProcedure
        Case 3228     'Invalid CollatingSequence entry in [Paradox ISAM] section in INI file.
            GetAction = ExitProcedure
        Case 3229     'Btrieve--can't change field.
            GetAction = ExitProcedure
        Case 3230     'Out-of-date Paradox lock file.
            GetAction = ExitProcedure
        Case 3231     'ODBC--field would be too long; data truncated.
            GetAction = ExitProcedure
        Case 3232     'ODBC--couldn't create table.
            GetAction = ExitProcedure
        Case 3233     'ODBC--incorrect driver version.
            GetAction = ExitProcedure
        Case 3234     'ODBC--remote query timeout expired.
            GetAction = ExitProcedure
        Case 3235     'ODBC--data type not supported on server.
            GetAction = ExitProcedure
        Case 3236     'ODBC--encountered unexpected Null value.
            GetAction = ExitProcedure
        Case 3237     'ODBC--unexpected type.
            GetAction = ExitProcedure
        Case 3238     'ODBC--data out of range.
            GetAction = ExitProcedure
        Case 3239     'Too many active users.
            GetAction = ExitProcedure
        Case 3240     'Btrieve--missing WBTRCALL.DLL.
            GetAction = ExitProcedure
        Case 3241     'Btrieve--out of resources.
            GetAction = ExitProcedure
        Case 3242     'Invalid reference in SELECT statement.
            GetAction = ExitProcedure
        Case 3243     'None of the import field names match fields in the appended table.
            GetAction = ExitProcedure
        Case 3244     'Can't import password-protected spreadsheet.
            GetAction = ExitProcedure
        Case 3245     'Couldn't parse field names from first row of import table.
            GetAction = ExitProcedure
        Case 3246     'Operation not supported in transactions.
            GetAction = ExitProcedure
        Case 3247     'ODBC--linked table definition has changed.
            GetAction = ExitProcedure
        Case 3248     'Invalid NetworkAccess entry in INI file.
            GetAction = ExitProcedure
        Case 3249     'Invalid PageTimeout entry in INI file.
            GetAction = ExitProcedure
        Case 3250     'Couldn't build key.
            GetAction = ExitProcedure
        Case 3251     'Feature not available.
            GetAction = ExitProcedure
        Case 3252     'Illegal reentrancy during query execution.
            GetAction = ExitProcedure
        Case 3254     'ODBC--Can't lock all records.
            GetAction = ExitProcedure
        Case 3255     'ODBC--Can't change connect string parameter.
            GetAction = ExitProcedure
        Case 3256     'Index file not found.
            GetAction = ExitProcedure
        Case 3257     'Syntax error in WITH OWNERACCESS OPTION declaration.
            GetAction = ExitProcedure
        Case 3258     'Query contains ambiguous (outer) joins.
            GetAction = ExitProcedure
        Case 3259     'Invalid field data type.
            GetAction = ExitProcedure
        Case 3260     'Couldn't update; currently locked by user ' 2' on machine '|1'.
            GetAction = ExitProcedure
        Case 3261     '
            GetAction = ExitProcedure
        Case 3262     '
            GetAction = ExitProcedure
        Case 3263     'Invalid database object.
            GetAction = ExitProcedure
        Case 3264     'No fields defined - cannot append table.
            GetAction = ExitProcedure
        Case 3266     'Append illegal - Field is part of a TableDefs collection.
            GetAction = ExitProcedure
        Case 3267     'Property value only valid when Field is part of a recordset.
            GetAction = ExitProcedure
        Case 3268     'Cannot set the property of an object which is part of a Database object.
            GetAction = ExitProcedure
        Case 3269     'Append illegal - Index is part of a TableDefs collection.
            GetAction = ExitProcedure
        Case 3270     'Property not found.
            GetAction = ExitProcedure
        Case 3271     'Invalid property value.
            GetAction = ExitProcedure
        Case 3272     'Object is not an array.
            GetAction = ExitProcedure
        Case 3273     'Method not applicable for this object.
            GetAction = ExitProcedure
        Case 3274     'External table isn't in the expected format.
            GetAction = ExitProcedure
        Case 3275     'Unexpected error from external database driver ( ).
            GetAction = ExitProcedure
        Case 3276     'Invalid database ID.
            GetAction = ExitProcedure
        Case 3277     'Can't have more than 10 fields in an index.
            GetAction = ExitProcedure
        Case 3278     'Database engine has not been initialized.
            GetAction = ExitProcedure
        Case 3279     'Database engine has already been initialized.
            GetAction = ExitProcedure
        Case 3280     'Can't delete a field that is part of an index.
            GetAction = ExitProcedure
        Case 3281     'Can't delete an index that is used in a relationship.
            GetAction = ExitProcedure
        Case 3282     'Can't perform operation on a nontable.
            GetAction = ExitProcedure
        Case 3283     'Primary key already exists.
            GetAction = ExitProcedure
        Case 3284     'Index already exists.
            GetAction = ExitProcedure
        Case 3285     'Invalid index definition.
            GetAction = ExitProcedure
        Case 3286     'Invalid type for Memo field.
            GetAction = ExitProcedure
        Case 3287     'Can't create index on Memo field or Long Binary field.
            GetAction = ExitProcedure
        Case 3288     'Invalid ODBC driver.
            GetAction = ExitProcedure
        Case 3289     'Paradox: No primary index.
            GetAction = ExitProcedure
        Case 3290     'Syntax error.
            GetAction = ExitProcedure
        Case 3291     'Syntax error in CREATE TABLE statement.
            GetAction = ExitProcedure
        Case 3292     'Syntax error in CREATE INDEX statement.
            GetAction = ExitProcedure
        Case 3293     'Syntax error in column definition.
            GetAction = ExitProcedure
        Case 3294     'Syntax error in ALTER TABLE statement.
            GetAction = ExitProcedure
        Case 3295     'Syntax error in DROP INDEX statement.
            GetAction = ExitProcedure
        Case 3296     'Syntax error in DROP statement.
            GetAction = ExitProcedure
        Case 3297     'Operation not supported in version 1.1
            GetAction = ExitProcedure
        Case 3298     'Couldn't import. No records found or all records contained errors.
            GetAction = ExitProcedure
        Case 3299     'Several tables exist with that name; please specify owner, as in
            GetAction = ExitProcedure
        Case 3265           'Item could not be found in the corresponding selection
            GetAction = ExitForm
        Case 3356           'Database is opened exclusively by another user
            GetAction = ExitProcedure
        Case 30006          'Unable to Allocate Memory
            GetAction = ExitForm
        Case 31001          'Out of memory
            GetAction = ExitForm
        Case 31004          'No object
            GetAction = ExitForm
        Case 31018          'Class is not set
            GetAction = ExitForm
        Case 31027          'Unable to activate object
            GetAction = ExitForm
        Case 31032          'Unable to create embedded object
            GetAction = ExitForm
        Case 31036          'Error saving to file
            GetAction = ExitForm
        Case 31037          'Error loading from file
            GetAction = ExitForm
'///////////////////////////////////////////////// !!CRYSTAL REPORT ERRORS!! \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Case -2147189547    'Too Many Connections to Crystal Reports
            GetAction = ExitForm
        Case -2147192184    'Automation Error
            GetAction = ExitForm
        Case -2147417848    'Cant initialize the objReport object
            GetAction = ExitForm
        Case -2147192153    'Missing Driver
            GetAction = ExitForm
        Case -2147217873    'Key Violation
            GetAction = ExitProcedure
        Case -2147467259    'Database is Locked or not found
            GetAction = ExitForm
        Case Else           'Default Action
            GetAction = ResumeNext
    End Select
End Function

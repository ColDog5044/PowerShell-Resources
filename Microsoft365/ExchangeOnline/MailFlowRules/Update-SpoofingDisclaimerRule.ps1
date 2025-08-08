#Requires -Version 5.1

<#
.SYNOPSIS
    Updates Exchange Online mail flow rule to include user display names and common name variations for spoofing disclaimer.

.DESCRIPTION
    This script connects to Exchange Online, retrieves all user display names, generates common name alterations (like Abe/Abraham),
    and updates the "Apply Disclaimer for Potential Spoofs" transport rule with the complete list of names to monitor for potential spoofing attempts.

.PARAMETER RuleName
    The name of the transport rule to update. Defaults to "Apply Disclaimer for Potential Spoofs".

.INPUTS
    RuleName - The name of the Exchange Online transport rule to update

.OUTPUTS
    Updates Exchange Online transport rule with user display names

.NOTES
    Version:          1.0
    Author:           Collin Laney
    Creation Date:    2025-08-07

    Exit Codes:
    Code 0 = Success
    Code 1 = Failure
    Code 2 = Script Error

    Requirements:
    - ExchangeOnlineManagement PowerShell module
    - Exchange Online administrator permissions

.EXAMPLE
    .\Update-SpoofingDisclaimerRule.ps1
    Runs the script with default parameters.

.EXAMPLE
    .\Update-SpoofingDisclaimerRule.ps1 -RuleName "Custom Spoofing Rule"
    Runs the script with a custom rule name.

.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/set-transportrule
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [string]$RuleName = "Apply Disclaimer for Potential Spoofs"
)

begin {
    # Initialize variables and functions here
    
    # Define names to exclude from the spoofing rule
    $excludeNames = @(
        "Mailbox",
        "DMARC",
        "Helpdesk")
    
    function Test-ModuleInstalled {
        param (
            [string]$ModuleName
        )
        
        try {
            $module = Get-Module -ListAvailable -Name $ModuleName
            if ($module) {
                Write-Verbose "$ModuleName module is installed (Version: $($module.Version -join ', '))"
                return $true
            }
            else {
                Write-Verbose "$ModuleName module is not installed"
                return $false
            }
        }
        catch {
            Write-Verbose "Error checking for $ModuleName module: $($_.Exception.Message)"
            return $false
        }
    }
    
    function Install-RequiredModule {
        param (
            [string]$ModuleName
        )
        
        try {
            Write-Host "Installing $ModuleName module..." -ForegroundColor Yellow
            Install-Module -Name $ModuleName -Force -AllowClobber -Scope CurrentUser
            Write-Host "$ModuleName module installed successfully" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Error "Failed to install $ModuleName module: $($_.Exception.Message)"
            return $false
        }
    }
    
    function Test-ExchangeOnlineConnection {
        try {
            $session = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
            if ($session) {
                Write-Verbose "Exchange Online session is already established"
                return $true
            }
            else {
                Write-Verbose "No active Exchange Online session found"
                return $false
            }
        }
        catch {
            Write-Verbose "Error checking Exchange Online connection: $($_.Exception.Message)"
            return $false
        }
    }

    function Get-NameAlterations {
        return @{
            "Abe"         = "Abraham"
            "Abraham"     = "Abe"
            "Abby"        = "Abigail"
            "Abigail"     = "Abby"
            "Adam"        = "Adan"
            "Adan"        = "Adam"
            "Addie"       = "Adelaide"
            "Adelaide"    = "Addie"
            "Adrian"      = "Ade"
            "Ade"         = "Adrian"
            "Al"          = "Albert"
            "Albert"      = "Al"
            "Alec"        = "Alexander"
            "Alex"        = "Alexander"
            "Alexander"   = "Alex"
            "Ali"         = "Alice"
            "Alice"       = "Ali"
            "Allie"       = "Alice"
            "Amanda"      = "Mandy"
            "Mandy"       = "Amanda"
            "Amy"         = "Amelia"
            "Amelia"      = "Amy"
            "Andre"       = "Andrew"
            "Andy"        = "Andrew"
            "Andrew"      = "Andy"
            "Angie"       = "Angela"
            "Angela"      = "Angie"
            "Ann"         = "Anna"
            "Anna"        = "Ann"
            "Annie"       = "Ann"
            "Anthony"     = "Tony"
            "Tony"        = "Anthony"
            "Artie"       = "Arthur"
            "Arthur"      = "Artie"
            "Art"         = "Arthur"
            "Ashley"      = "Ash"
            "Ash"         = "Ashley"
            "Austin"      = "Gus"
            "Gus"         = "Austin"
            "Barb"        = "Barbara"
            "Barbara"     = "Barb"
            "Barry"       = "Bernard"
            "Bernard"     = "Barry"
            "Bart"        = "Bartholomew"
            "Bartholomew" = "Bart"
            "Bea"         = "Beatrice"
            "Beatrice"    = "Bea"
            "Becky"       = "Rebecca"
            "Rebecca"     = "Becky"
            "Ben"         = "Benjamin"
            "Benjamin"    = "Ben"
            "Benny"       = "Benjamin"
            "Bernie"      = "Bernard"
            "Beth"        = "Elizabeth"
            "Elizabeth"   = "Beth"
            "Betty"       = "Elizabeth"
            "Beverly"     = "Bev"
            "Bev"         = "Beverly"
            "Bill"        = "William"
            "Billy"       = "William"
            "Bob"         = "Robert"
            "Bobby"       = "Robert"
            "Brad"        = "Bradley"
            "Bradford"    = "Brad"
            "Bradley"     = "Brad"
            "Brandon"     = "Bran"
            "Bran"        = "Brandon"
            "Brenda"      = "Brendan"
            "Brendan"     = "Brenda"
            "Brian"       = "Bryan"
            "Bryan"       = "Brian"
            "Brittany"    = "Britt"
            "Britt"       = "Brittany"
            "Bruce"       = "Bruno"
            "Bruno"       = "Bruce"
            "Cal"         = "Calvin"
            "Calvin"      = "Cal"
            "Cameron"     = "Cam"
            "Cam"         = "Cameron"
            "Carl"        = "Carlos"
            "Carlos"      = "Carl"
            "Carol"       = "Caroline"
            "Caroline"    = "Carol"
            "Carrie"      = "Caroline"
            "Casey"       = "Cassandra"
            "Cassandra"   = "Casey"
            "Cathy"       = "Catherine"
            "Catherine"   = "Cathy"
            "Chad"        = "Charles"
            "Charlie"     = "Charles"
            "Charles"     = "Charlie"
            "Charlotte"   = "Lottie"
            "Lottie"      = "Charlotte"
            "Chloe"       = "Cloe"
            "Cloe"        = "Chloe"
            "Chris"       = "Christopher"
            "Christopher" = "Chris"
            "Christine"   = "Christina"
            "Christina"   = "Christine"
            "Chuck"       = "Charles"
            "Cindy"       = "Cynthia"
            "Cynthia"     = "Cindy"
            "Claire"      = "Clara"
            "Clara"       = "Claire"
            "Clarence"    = "Clare"
            "Clare"       = "Clarence"
            "Colleen"     = "Cole"
            "Cole"        = "Colleen"
            "Connor"      = "Con"
            "Con"         = "Connor"
            "Craig"       = "Greg"
            "Curtis"      = "Curt"
            "Curt"        = "Curtis"
            "Dale"        = "Dallas"
            "Dallas"      = "Dale"
            "Dan"         = "Daniel"
            "Daniel"      = "Dan"
            "Danny"       = "Daniel"
            "Darren"      = "Dare"
            "Dare"        = "Darren"
            "Dave"        = "David"
            "David"       = "Dave"
            "Davy"        = "David"
            "Dean"        = "Deane"
            "Deane"       = "Dean"
            "Debbie"      = "Deborah"
            "Deborah"     = "Debbie"
            "Deb"         = "Deborah"
            "Dennis"      = "Denny"
            "Denny"       = "Dennis"
            "Derek"       = "Derrick"
            "Derrick"     = "Derek"
            "Diana"       = "Di"
            "Di"          = "Diana"
            "Dolores"     = "Dolly"
            "Dolly"       = "Dolores"
            "Don"         = "Donald"
            "Donald"      = "Don"
            "Donnie"      = "Donald"
            "Dorothy"     = "Dot"
            "Dot"         = "Dorothy"
            "Doug"        = "Douglas"
            "Douglas"     = "Doug"
            "Drew"        = "Andrew"
            "Earl"        = "Early"
            "Early"       = "Earl"
            "Ed"          = "Edward"
            "Edward"      = "Ed"
            "Eddie"       = "Edward"
            "Eileen"      = "Ellie"
            "Eleanor"     = "Ellie"
            "Ellie"       = "Eleanor"
            "Emily"       = "Em"
            "Em"          = "Emily"
            "Emma"        = "Emmy"
            "Emmy"        = "Emma"
            "Eric"        = "Erik"
            "Erik"        = "Eric"
            "Ernest"      = "Ernie"
            "Ernie"       = "Ernest"
            "Eugene"      = "Gene"
            "Gene"        = "Eugene"
            "Evan"        = "Ev"
            "Ev"          = "Evan"
            "Faith"       = "Faye"
            "Faye"        = "Faith"
            "Felix"       = "Phil"
            "Florence"    = "Flo"
            "Flo"         = "Florence"
            "Frances"     = "Fran"
            "Fran"        = "Frances"
            "Francis"     = "Frank"
            "Frank"       = "Francis"
            "Franklin"    = "Frank"
            "Fred"        = "Frederick"
            "Frederick"   = "Fred"
            "Freddie"     = "Frederick"
            "Gabriel"     = "Gabe"
            "Gabe"        = "Gabriel"
            "Gary"        = "Gareth"
            "Gareth"      = "Gary"
            "George"      = "Georg"
            "Georg"       = "George"
            "Georgie"     = "George"
            "Gerald"      = "Jerry"
            "Jerry"       = "Gerald"
            "Gerard"      = "Gerry"
            "Gerry"       = "Gerard"
            "Gilbert"     = "Gil"
            "Gil"         = "Gilbert"
            "Grace"       = "Gracie"
            "Gracie"      = "Grace"
            "Graham"      = "Gray"
            "Gray"        = "Graham"
            "Gregory"     = "Greg"
            "Greg"        = "Gregory"
            "Harold"      = "Harry"
            "Harry"       = "Harold"
            "Harvey"      = "Harv"
            "Harv"        = "Harvey"
            "Heather"     = "Heath"
            "Heath"       = "Heather"
            "Helen"       = "Lena"
            "Lena"        = "Helen"
            "Henry"       = "Hank"
            "Hank"        = "Henry"
            "Herbert"     = "Herb"
            "Herb"        = "Herbert"
            "Hope"        = "Hopie"
            "Hopie"       = "Hope"
            "Howard"      = "Howie"
            "Howie"       = "Howard"
            "Ian"         = "Iain"
            "Iain"        = "Ian"
            "Isaac"       = "Ike"
            "Ike"         = "Isaac"
            "Isabel"      = "Izzy"
            "Izzy"        = "Isabel"
            "Jackson"     = "Jack"
            "Jacob"       = "Jake"
            "Jake"        = "Jacob"
            "James"       = "Jim"
            "Jim"         = "James"
            "Jamie"       = "James"
            "Jimmy"       = "James"
            "Jan"         = "Janet"
            "Janet"       = "Jan"
            "Jane"        = "Janie"
            "Janie"       = "Jane"
            "Jason"       = "Jay"
            "Jay"         = "Jason"
            "Jean"        = "Jeanie"
            "Jeanie"      = "Jean"
            "Jeff"        = "Jeffrey"
            "Jeffrey"     = "Jeff"
            "Jen"         = "Jennifer"
            "Jennifer"    = "Jen"
            "Jenny"       = "Jennifer"
            "Jeremy"      = "Jerry"
            "Jessica"     = "Jess"
            "Jess"        = "Jessica"
            "Jessie"      = "Jessica"
            "Jill"        = "Jillian"
            "Jillian"     = "Jill"
            "Joan"        = "Joanie"
            "Joanie"      = "Joan"
            "Joanna"      = "Jo"
            "Jo"          = "Joanna"
            "Joe"         = "Joseph"
            "Joseph"      = "Joe"
            "Joey"        = "Joseph"
            "John"        = "Johnny"
            "Johnny"      = "John"
            "Jon"         = "Jonathan"
            "Jonathan"    = "Jon"
            "Jordan"      = "Jordy"
            "Jordy"       = "Jordan"
            "Jose"        = "Joseph"
            "Josh"        = "Joshua"
            "Joshua"      = "Josh"
            "Joyce"       = "Joy"
            "Joy"         = "Joyce"
            "Judith"      = "Judy"
            "Judy"        = "Judith"
            "Julia"       = "Julie"
            "Julie"       = "Julia"
            "Justin"      = "Just"
            "Just"        = "Justin"
            "Karen"       = "Kay"
            "Kay"         = "Karen"
            "Kate"        = "Katherine"
            "Katherine"   = "Kate"
            "Katie"       = "Katherine"
            "Keith"       = "Key"
            "Key"         = "Keith"
            "Kelly"       = "Kel"
            "Kel"         = "Kelly"
            "Ken"         = "Kenneth"
            "Kenneth"     = "Ken"
            "Kenny"       = "Kenneth"
            "Kevin"       = "Kev"
            "Kev"         = "Kevin"
            "Kim"         = "Kimberly"
            "Kimberly"    = "Kim"
            "Kyle"        = "Ky"
            "Ky"          = "Kyle"
            "Larry"       = "Lawrence"
            "Lawrence"    = "Larry"
            "Laura"       = "Laurie"
            "Laurie"      = "Laura"
            "Lee"         = "Leslie"
            "Leslie"      = "Lee"
            "Leonard"     = "Leo"
            "Leo"         = "Leonard"
            "Lewis"       = "Lou"
            "Linda"       = "Lindy"
            "Lindy"       = "Linda"
            "Lisa"        = "Elisabeth"
            "Elisabeth"   = "Lisa"
            "Liz"         = "Elizabeth"
            "Logan"       = "Lo"
            "Lo"          = "Logan"
            "Lori"        = "Lorraine"
            "Lorraine"    = "Lori"
            "Lou"         = "Louis"
            "Louis"       = "Lou"
            "Louise"      = "Lou"
            "Lucas"       = "Luke"
            "Luke"        = "Lucas"
            "Lucille"     = "Lucy"
            "Lucy"        = "Lucille"
            "Lynn"        = "Lynne"
            "Lynne"       = "Lynn"
            "Madison"     = "Maddie"
            "Maddie"      = "Madison"
            "Margaret"    = "Meg"
            "Meg"         = "Margaret"
            "Maggie"      = "Margaret"
            "Maria"       = "Mary"
            "Mary"        = "Maria"
            "Marie"       = "Mary"
            "Mark"        = "Marcus"
            "Marcus"      = "Mark"
            "Martin"      = "Marty"
            "Marty"       = "Martin"
            "Mason"       = "Mase"
            "Mase"        = "Mason"
            "Matt"        = "Matthew"
            "Matthew"     = "Matt"
            "Maurice"     = "Mo"
            "Mo"          = "Maurice"
            "Max"         = "Maxwell"
            "Maxwell"     = "Max"
            "Maxine"      = "Max"
            "Megan"       = "Meg"
            "Melanie"     = "Mel"
            "Mel"         = "Melanie"
            "Melissa"     = "Missy"
            "Missy"       = "Melissa"
            "Michael"     = "Mike"
            "Mike"        = "Michael"
            "Mickey"      = "Michael"
            "Michelle"    = "Shelly"
            "Shelly"      = "Michelle"
            "Miranda"     = "Randy"
            "Mitchell"    = "Mitch"
            "Mitch"       = "Mitchell"
            "Morgan"      = "Mo"
            "Nancy"       = "Ann"
            "Natalie"     = "Nat"
            "Nathan"      = "Nat"
            "Nat"         = "Nathan"
            "Nate"        = "Nathan"
            "Neil"        = "Neal"
            "Neal"        = "Neil"
            "Nicholas"    = "Nick"
            "Nick"        = "Nicholas"
            "Nicole"      = "Nicky"
            "Nicky"       = "Nicole"
            "Noah"        = "No"
            "No"          = "Noah"
            "Norman"      = "Norm"
            "Norm"        = "Norman"
            "Oliver"      = "Ollie"
            "Ollie"       = "Oliver"
            "Oscar"       = "Oz"
            "Oz"          = "Oscar"
            "Pamela"      = "Pam"
            "Pam"         = "Pamela"
            "Patricia"    = "Patty"
            "Patty"       = "Patricia"
            "Patrick"     = "Pat"
            "Pat"         = "Patrick"
            "Paul"        = "Paulo"
            "Paulo"       = "Paul"
            "Pauline"     = "Polly"
            "Polly"       = "Pauline"
            "Peter"       = "Pete"
            "Pete"        = "Peter"
            "Philip"      = "Phil"
            "Phil"        = "Philip"
            "Phillip"     = "Phil"
            "Preston"     = "Press"
            "Press"       = "Preston"
            "Rachel"      = "Ray"
            "Randall"     = "Randy"
            "Randy"       = "Randall"
            "Raymond"     = "Ray"
            "Ray"         = "Raymond"
            "Reginald"    = "Reggie"
            "Reggie"      = "Reginald"
            "Richard"     = "Rick"
            "Rick"        = "Richard"
            "Ricky"       = "Richard"
            "Robert"      = "Rob"
            "Rob"         = "Robert"
            "Robbie"      = "Robert"
            "Robin"       = "Rob"
            "Roger"       = "Rog"
            "Rog"         = "Roger"
            "Ronald"      = "Ron"
            "Ron"         = "Ronald"
            "Ronnie"      = "Ronald"
            "Rose"        = "Rosie"
            "Rosie"       = "Rose"
            "Russell"     = "Russ"
            "Russ"        = "Russell"
            "Ryan"        = "Ry"
            "Ry"          = "Ryan"
            "Samantha"    = "Sam"
            "Samuel"      = "Sam"
            "Sam"         = "Samuel"
            "Sammy"       = "Samuel"
            "Sandra"      = "Sandy"
            "Sandy"       = "Sandra"
            "Sara"        = "Sarah"
            "Sarah"       = "Sara"
            "Scott"       = "Scot"
            "Scot"        = "Scott"
            "Scotty"      = "Scott"
            "Sebastian"   = "Seb"
            "Seb"         = "Sebastian"
            "Shane"       = "Shay"
            "Shay"        = "Shane"
            "Sharon"      = "Shari"
            "Shari"       = "Sharon"
            "Stephanie"   = "Steph"
            "Steph"       = "Stephanie"
            "Stephen"     = "Steve"
            "Steve"       = "Stephen"
            "Steven"      = "Steve"
            "Stuart"      = "Stu"
            "Stu"         = "Stuart"
            "Susan"       = "Sue"
            "Sue"         = "Susan"
            "Susie"       = "Susan"
            "Taylor"      = "Tay"
            "Tay"         = "Taylor"
            "Terence"     = "Terry"
            "Terry"       = "Terence"
            "Teresa"      = "Terry"
            "Theodore"    = "Ted"
            "Ted"         = "Theodore"
            "Theresa"     = "Terry"
            "Thomas"      = "Tom"
            "Tom"         = "Thomas"
            "Tommy"       = "Thomas"
            "Timothy"     = "Tim"
            "Tim"         = "Timothy"
            "Timmy"       = "Timothy"
            "Todd"        = "Tod"
            "Tod"         = "Todd"
            "Tracy"       = "Trace"
            "Trace"       = "Tracy"
            "Trevor"      = "Trev"
            "Trev"        = "Trevor"
            "Tyler"       = "Ty"
            "Ty"          = "Tyler"
            "Valerie"     = "Val"
            "Val"         = "Valerie"
            "Vanessa"     = "Nessa"
            "Nessa"       = "Vanessa"
            "Victoria"    = "Vicky"
            "Vicky"       = "Victoria"
            "Vincent"     = "Vinny"
            "Vinny"       = "Vincent"
            "Virginia"    = "Ginny"
            "Ginny"       = "Virginia"
            "Walter"      = "Walt"
            "Walt"        = "Walter"
            "Wayne"       = "Way"
            "Way"         = "Wayne"
            "Wesley"      = "Wes"
            "Wes"         = "Wesley"
            "William"     = "Will"
            "Will"        = "William"
            "Willie"      = "William"
            "Zachary"     = "Zach"
            "Zach"        = "Zachary"
            "Zoe"         = "Zoey"
            "Zoey"        = "Zoe"
            # Add more name variations or adjust as needed
        }
    }

    function New-NameVariations {
        param (
            [string[]]$Users,
            [hashtable]$NameAlterations
        )
        
        $allNames = @()
        foreach ($user in $Users) {
            $allNames += $user
            
            # Split the user's name into parts
            $nameParts = $user -split " "
            
            foreach ($key in $NameAlterations.Keys) {
                # Check each part of the name for matches
                foreach ($part in $nameParts) {
                    if ($part -eq $key) {
                        # Create alternate version with the same last name structure
                        $alternateUser = $user -replace [regex]::Escape($key), $NameAlterations[$key]
                        $allNames += $alternateUser
                    }
                }
            }
        }
        
        # Remove duplicates and return
        return ($allNames | Sort-Object -Unique)
    }
}

process {
    try {
        # Check if ExchangeOnlineManagement module is installed
        if (-not (Test-ModuleInstalled -ModuleName "ExchangeOnlineManagement")) {
            Write-Host "ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
            if (-not (Install-RequiredModule -ModuleName "ExchangeOnlineManagement")) {
                Write-Error "Failed to install ExchangeOnlineManagement module. Cannot continue."
                exit 2
            }
        }
        
        # Import the module if not already loaded
        if (-not (Get-Module -Name "ExchangeOnlineManagement")) {
            Write-Host "Importing ExchangeOnlineManagement module..." -ForegroundColor Yellow
            Import-Module ExchangeOnlineManagement -Force
        }

        # Check if Exchange Online connection exists, if not connect
        if (-not (Test-ExchangeOnlineConnection)) {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
            Connect-ExchangeOnline
        }

        # Get all user display names
        Write-Host "Retrieving user display names..." -ForegroundColor Yellow
        $users = Get-User -ResultSize Unlimited | 
        Where-Object { $_.DisplayName -notlike "*Mailbox*" -and $_.DisplayName -notin $excludeNames } |
        Select-Object -ExpandProperty DisplayName

        Write-Verbose "Found $($users.Count) users to process"

        # Define common name alterations
        $nameAlterations = Get-NameAlterations

        # Create a list of names including alterations
        Write-Host "Generating name variations..." -ForegroundColor Yellow
        $uniqueNames = New-NameVariations -Users $users -NameAlterations $nameAlterations

        Write-Verbose "Generated $($uniqueNames.Count) unique name variations"

        # Update the mail flow rule
        Write-Host "Updating transport rule '$RuleName'..." -ForegroundColor Yellow
        Set-TransportRule -Identity $RuleName -HeaderMatchesMessageHeader "From" -HeaderMatchesPatterns $uniqueNames

        Write-Host "Successfully updated transport rule with $($uniqueNames.Count) name patterns" -ForegroundColor Green
        exit 0
    }
    catch {
        Write-Error "Script execution failed: $($_.Exception.Message)"
        exit 1
    }
}

end {
    # Clean up code here
    try {
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Yellow
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Script completed successfully" -ForegroundColor Green
    }
    catch {
        Write-Warning "Error during cleanup: $($_.Exception.Message)"
        exit 2
    }
}
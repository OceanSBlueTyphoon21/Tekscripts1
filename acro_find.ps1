# PS Script: acro_find
# Description: Finds a Tektronix Business acronym and defines the abbreviation 


# import Modules
Import-Module ImportExcel

# variables
$acronym_list = Import-Excel "C:\Users\u619233\OneDrive - Fortive\tekImmersionPlan.xlsx" -WorksheetName Acronyms
$CurrentAcronym = ""


# WHILE loop, reprompts user for any Acronym unless the command "close_acro" is entered
while($true)
{
    # prompt user for acronym
    $User_acronym = Read-Host 'Enter a Tek Acronym '

    if ($User_acronym -ieq 'close_acro')   # Determine if the User_acronym is "close_acro"
    {
        break   # If true, break out of the reprompting while loop
    }

    else  # if False, default to determining if the User_acronym is in the Tektronix Business Acronym Excel Sheet (TBAES)
    {
        for ($i=0; $i -le $acronym_list.Length-1; $i++) # Loop through TBAES for the acronym
        {
            $CurrentAcronym = $acronym_list[$i].Acronym  # set the TBAES acronym at index i to equal CurrentAcronym
            if($User_acronym -ieq $CurrentAcronym)       # Compare the CurrentAcronym (from TBAES) to User_acronym
            {
                $acronym_list[$i].Description        # If True, the Acronym was found in the TBAES. Write to console
                Write-Output "`n"                    # Newline in console
                break   # break out of for-loop
            }
        }

        if ($i -ge $acronym_list.Length-1)  # Determine if we have gone through the entire TBAES list
        {
            Write-Output "No Description found for Acronym`n"  # If True, return the error message
        }
    }
}
clear  # Clear the console information
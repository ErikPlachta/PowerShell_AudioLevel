<# 
    Author: Erik Plachta
    Date: 11/09/2021
    Purpose: Manage Windows Audio Levels With PowerShell
    Summary: Using PowerShell to create a COM Object of the type Windows Shell. Then running Windows Shell function function SendKeys() with the parameters `[char]173`, `[char]174`, or `[char]175`.
#>

<# % FUNCTION - START #>
function Set-Speaker($audio_Level){
    <# Loop to adjust audio level based on volume

        Order of Operations    
            1. Create Windows Shell object
            2. Divides $audio_Level by 50% because command uses 2% by default, but we are asssuming 1%.        
            3. Sets Volume Level to 0
            4. Increments UP to Level based on provided INT $audio_Level

        Example: ( Remember incrementing by 1% )
            
            #Sets volume to 60%
            Set-Speaker -Volume 60

            #Sets volume to 80%
            Set-Speaker -Volume 80

            #Sets volume to 100%
            Set-Speaker -Volume 100
    #>

    #1. Create Shell  Object
    $wshShell = new-object -com wscript.shell;

     #2. Make each level 1 %. by Default assumes 2%, want to be 1 so can do 0 - 100% level
     $volume_Level = $audio_Level / 2
     
     ##3. 50 time loop
     1..50 | % {
         ##Set Volume Level to ZERO ( 50 = MAX or 100% level )
         $wshShell.SendKeys([char]174)
     }
     
     ##4. Loop from 1 to volume level provide
     1..$volume_Level | % {
         ## Adjust volume level UP for by one for each level loop
         $wshShell.SendKeys([char]175)
     }

}
<# 
    Author: Erik Plachta
    Date: 11/09/2021
    Purpose: Manage Windows Audio Volume Levels With PowerShella
    Summary: Using PowerShell to create a COM Object of the type Windows Shell. Then running Windows Shell function function SendKeys() with the parameters `[char]173`, `[char]174`, or `[char]175`.
#>

<# % FUNCTION - START #>
function set_AudioLevel($audioLevel){
    <# Loop to adjust audio level based on volume

        Order of Operations    
            1. Create Windows Shell object
            2. Divides $audio_Level by 50% because command uses 2% by default, but we are assuming 1%.        
            3. Sets Volume Level to 0
            4. Increments UP to Level based on provided INT $audio_Level

        Example: ( Remember incrementing by 1% )
            
            #Sets volume to 60%
            set_AudioLevel -Volume 60

            #Sets volume to 80%
            set_AudioLevel -Volume 80

            #Sets volume to 100%
            set_AudioLevel -Volume 100
    #>

    ## 1. Create Shell  Object
    $wshShell = new-object -com wscript.shell;

     ## 2. Sets Volume Level to 0
     1..50 | % { # 50 time loop
         ##Set Volume Level to ZERO ( 50 = MAX or 100% level )
         $wshShell.SendKeys([char]174)
     }
     
     ## 3. Make each level 1 %. by Default assumes 2%, want to be 1 so can do 0 - 100% level
     $audioLevel = $audioLevel / 2
     
     ## 4. Increments UP to Level based on provided INT $audio_Level
     1..$audioLevel | % { ## Loop from 1 to volume level provide
         ## Adjust volume level UP for by one for each level loop
         $wshShell.SendKeys([char]175)
     }

}

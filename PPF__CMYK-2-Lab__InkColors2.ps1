#====================================================================================================================
#====================================================================================================================
#====================================================================================================================

function produit_matrice($MA, $MB){    
    # J'arrive pas à utiliser cette fonction magnifique .. Il y a un genre de transtypage des
    #  arguments $MA et $MB ..., le .lenght n'est pas conforme ..  J'y arrive PAS on la garde mais 
    #  je Copie Colle ça dans RGB2XYZ a la vollée. 
    echo "MA from function : " 
    echo $MA;
    $MA = @(@(1,2,3), @(4,5,6), @(7,8,9))
    echo "MA from code :" 
    echo $MA
    $MAlines = $MA.Length
    $MAcols = $MA[0].Length
    echo "Lignes M1 = $MAlines"
    echo "Colones M1 = $MAcols"

    #$MB = @(@(10,13), @(11,14), @(12,15))
    $MBlines = $MB.Length
    $MBcols = $MB[0].Length
    echo "Lignes M2 = $MBlines"
    echo "Colones M2 = $MBcols"

    #$MA = $MA | ConvertFrom-Json
    echo "------------------"
    if ($MAcols -eq $MBlines){
        Write-output $MA
        $MF = New-Object 'object[,]' $MAlines,$MBcols

        For($i=0;$i -lt $MBcols;$i++) { 
            For($j=0;$j -lt $MAlines;$j++){
                Write-Output "L'emplacement MAT finale i,j est $($i,$j)" 
                # Calculer pour chaque emplacement : 
                # Somme de la multiplication de la ligne M1 avec la Col M2 
                $result = 0
                For($pi=0;$pi -lt $MAcols;$pi++){
                    $MAval=$MA[$j][$pi];
                    $MBval=$MB[$pi][$i]
                    echo "MAval= $MAval";
                    echo "MBval= $MBval";
                    $result = $result + ($MAval*$MBval)
                    echo "Tresult = $result"; 
                }
                $MF[$j,$i]=$result;
            }
        }clear
        #FINALLY SHOW MF : 
        echo $MF
        return $MF
    }else{
        echo "FALSE"
      #return false;
    }
}
#--------------------------------------------------------------------------------------------------------------------
function CMYK2RGB($CMYK){
    #echo "CMYK2RGB()"
    #echo $CMYK;
    
    [int]$C = $CMYK[0];
    [int]$M = $CMYK[1];
    [int]$Y = $CMYK[2];
    [int]$K = $CMYK[3];
    #echo "C= $C"
    #echo "M= $M"
    #echo "Y= $Y"
    #echo "K= $K"
    #echo "----"
    [int]$R = 255*(1-($C/100))*(1-($K/100))
    [int]$G = 255*(1-($M/100))*(1-($K/100))
    [int]$B = 255*(1-($Y/100))*(1-($K/100))
    #echo "R= $R"
    #echo "G= $G"
    #echo "B= $B"
    $textRGBColor = "[$R,$G,$B]"
    #echo $textRGBColor
    return $RGBColor = $textRGBColor | ConvertFrom-Json
}
#--------------------------------------------------------------------------------------------------------------------
function gamma2Linear($V){
    if ($V -lt 0.04045){
        return $V/12.92
    }else{
        return [math]::Pow((($V+0.055)/1.055),2.4)
    }
}
function RGB2XYZ($RGBColor){
    #echo "RGB2XYZ()"
    #Avant ça on doit gérer le taux de Gamma .. 
    #$MA = @(@(0.4124564,0.3575761,0.1804375), 
    #        @(0.2126729,0.7151522,0.0721750), 
    #        @(0.0193339,0.1191920,0.9503041)); # MATRICE sRGB D65
    $MA = @(@(0.4360747,0.3850649,0.1430804), 
            @(0.2225045,0.7168786,0.0606169), 
            @(0.0139322,0.0971045,0.7141733)); # MATRICE sRGB D50
    $R=$RGBColor[0]/255
    $G=$RGBColor[1]/255
    $B=$RGBColor[2]/255
    #echo "R=$R"
    #echo "G=$G"
    #echo "B=$B"
    #--------------------- sRGB Gamma to Linear ==
    $r = gamma2linear($R)
    $g = gamma2linear($G)
    $b = gamma2linear($B)

    #-----------------------------------------------------------
    $MB = @(@($r), @($g), @($b));          # MATRICE Couleur RGB
    #-----------------------
    $MAlines = $MA.Length
    $MAcols = $MA[0].Length
    #echo "Lignes M1 = $MAlines"
    #echo "Colones M1 = $MAcols"
    #-----------------------
    $MBlines = $MB.Length
    $MBcols = $MB[0].Length
    #echo "Lignes M2 = $MBlines"
    #echo "Colones M2 = $MBcols"

    #$MA = $MA | ConvertFrom-Json
    #echo "------------------"
    if ($MAcols -eq $MBlines){
        #Write-output $MA
        $MF = New-Object 'object[,]' $MAlines,$MBcols

        For($i=0;$i -lt $MBcols;$i++) { 
            For($j=0;$j -lt $MAlines;$j++){
                #Write-Output "L'emplacement MAT finale i,j est $($i,$j)" 
                # Calculer pour chaque emplacement : 
                # Somme de la multiplication de la ligne M1 avec la Col M2 
                $result = 0
                For($pi=0;$pi -lt $MAcols;$pi++){
                    $MAval=$MA[$j][$pi];
                    $MBval=$MB[$pi][$i]
                    #echo "MAval= $MAval";
                    #echo "MBval= $MBval";
                    $result = $result + ($MAval*$MBval)
                    #echo "Tresult = $result"; 
                }
                $MF[$j,$i]=$result;
                #echo $result;
            }
        }
        #FINALLY SHOW MF : 
        #echo "---------------------"
        #echo $MF
    }else{
        echo "Impossible de multiplier ces matrices !"
        return false;
    }
    return $MF;
}
#--------------------------------------------------------------------------------------------------------------------
function XYZ2LAB($X,$Y,$Z,$Illum="D50"){
    $illums = @{ "D50" = @(0.9642,1.0000,0.8251);
                 "D55" = @(0.9568,1.0000,0.9214);
                 "D65" = @(0.9504,1.0000,1.0888);
                 "icc" = @(0.9642,1.0000,0.8249);
                 "A" = @(1.0985,1.0000,0.3558);
                 "C" = @(0.9807,1.0000,1.1822);
                 "E" = @(1.0000,1.0000,1.0000)}
    $Xi = $illums[$Illum][0]
    $Yi = $illums[$Illum][1]
    $Zi = $illums[$Illum][2]

    #echo "Xi=$Xi"
    #echo "Yi=$Yi"
    #echo "Zi=$Zi"
    #echo "X=$X"
    #echo "Y=$Y"
    #echo "Z=$Z"


    if ($X/$Xi -gt 0.008856){
        $Fx= [Math]::pow(($X/$Xi),(1/3))
    }else{
        $Fx= ((903.3*($X/$Xi))+16)/116
    }
    if ($Y/$Yi -gt 0.008856){
        $Fy= [Math]::pow(($Y/$Yi),(1/3))
    }else{
        $Fy = ((903.3*($Y/$Yi))+16)/116
    }
    if ($Z/$Zi -gt 0.008856){
        $Fz= [Math]::pow(($Z/$Zi),(1/3))
    }else{
        $Fz = ((903.3*($Z/$Zi))+16)/116
    }
    #echo "(X/Xi)^1/3=$Fx"
    #echo "(Y/Yi)^1/3=$Fy"
    #echo "(Z/Zi)^1/3=$Fz"

    $L = (116*$Fy)-16
    $a = 500*($Fx-$Fy)
    $b = 200*($Fy-$Fz)
    $LABColor = @($L,$a,$b)
    #echo $LabColor
    return $LabColor
}





#====================================================================================================================
#====================================================================================================================
#====================================================================================================================

#$ppf_file = "E:\PTTransfer\PPFIn\SM_75\6903-22501P-A1-REPRISE_r_.ppf"


$PPF_InFolder = "E:\PTTransfer\PPFIn\XL_106\"
$PPF_OutFolder = "E:\PTTransfer\PPFIn\XL106-1"



$PPFFiles = Get-ChildItem $PPF_InFolder
foreach ($PPFFile in $PPFFiles){
echo  "#====================================================================================================================
#====================================================================================================================
#===================================================================================================================="
    $ppf_file = $PPFFile.FullName
    $fName = $PPFFile.Name
    $OutFullName = "$PPF_OutFolder\$fName.ppf"
    [string]$AdmInkColorsCMYKfull = Get-Content $ppf_file | Select-String -Pattern "/CIP3AdmInkColorsCMYK" 

    echo " "
    echo "Found line in CIP3 file :   $AdmInkColorsCMYKfull"
    $AdmInkColorsCMYK = $AdmInkColorsCMYKfull.replace("/CIP3AdmInkColorsCMYK ", "").replace(" def", "")
    echo "Table of Colors to parse :  $AdmInkColorsCMYK"
    echo " "
    echo "====================================== PARSING & COMPUTE =="


    $CMYKColors = Select-String -InputObject $AdmInkColorsCMYK -AllMatches -Pattern "\[[0-9]{1,3} [0-9]{1,3} [0-9]{1,3} [0-9]{1,3}\]" 
    echo $CMYKColors
    #$CMYKColors.Matches

    #echo $CMYK_Colors
    $finalTextLAB="["
    foreach ($textCMYKColor in $CMYKColors.Matches){
        echo "------------------ Couleur ------------------"
        echo $textCMYKColor
        $textCMYKColor = $textCMYKColor.value
        $textCMYKColor = $textCMYKColor.replace(" ",",")
        echo $textCMYKColor
        $CMYKColor = $textCMYKColor | ConvertFrom-Json
        echo "--> Couleur Quadri : "
        echo $CMYKColor
        $RGBColor = CMYK2RGB($CMYKColor)
        echo "--> Couleur RGB : "
        echo $RGBColor
        $XYZColor = RGB2XYZ($RGBColor)
        echo "--> Couleur XYZ : "
        echo $XYZColor

        $X = $XYZColor[0]
        $Y = $XYZColor[1]
        $Z = $XYZColor[2]
        $LabColor = XYZ2LAB $X $Y $Z "D50"
        echo "--> Couleur Lab : "
        echo $LabColor
        $L = ([math]::round($LabColor[0]*100))/100
        $a = ([math]::round($LabColor[1]*100))/100
        $b = ([math]::round($LabColor[2]*100))/100
    
        $textLAB = "[$L $a $b]"
        $finalTextLAB="$finalTextLAB $textLAB"
    }
    $finalTextLAB="$finalTextLAB ]"
    #echo $finalTextLAB
    $final_Replacer = "/CIP3AdmInkColors $finalTextLAB def"
    echo " "
    echo "====================================== COMPUTED =="
    echo "replace  '$AdmInkColorsCMYKfull'" 
    echo "by       '$final_Replacer'"
    echo " "
    echo "====================================== REPLACE & SAVE =="
    (Get-Content $ppf_file).replace($AdmInkColorsCMYKfull, $final_Replacer) | Set-Content $OutFullName
}
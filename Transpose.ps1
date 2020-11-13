function Transpose-Data{
    param(
        [String[]]$Names,
        [Object[][]]$Data
    )
    for($i = 0;; ++$i){
        $Props = [ordered]@{}
        for($j = 0; $j -lt $Data.Length; ++$j){
            if($i -lt $Data[$j].Length){
                $Props.Add($Names[$j], $Data[$j][$i])
            }
        }
        if(!$Props.get_Count()){
            break
        }
        [PSCustomObject]$Props
    }
}

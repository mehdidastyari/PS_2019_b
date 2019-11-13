function add
{
    write-host ( $args[0] + $args[1] )
}

add 10 20
$no1 = 10
$no2 = 20

add $no1 $no2
add (get-date).Year 1

add 1 (Get-Item * | 
    where { $_ -is [io.fileinfo] } | 
    sort length  | 
    select -last 1 ).lastwritetime.year



$TEMPLATES = ".\vorlagen.csv"


$Table = Import-CSV -Path "$TEMPLATES" -Delimiter ";" -Encoding "UTF8" -Header A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z

#foreach($Row in $Table) {
#		write-output "$($Row.A)"
#
#
#
#
#
#
#
#	}




#echo $Table.A


#foreach($i in $Table.A)
#{
#
#echo "go" $i
#
#}


echo $Table.A[0] # Template



echo "content"


$lengthstop = $Table.A.Length - 1
foreach($i in $Table.A[1..$lengthstop])
{

echo -join("ITEM",$i)

}
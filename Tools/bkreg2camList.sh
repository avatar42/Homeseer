rm -f cams.csv
echo  host,shortname,dsname,fullxres,fullyres,HS3ref,vcodec,folder,enabled,group > cams.csv
set -x
for host in BlueIris2 BlueIris3 BlueIris4 BlueIris5
do
	iconv -f UTF-16LE -t UTF-8  ${host}.reg | grep -E "\"enabled\"|\"group\"|shortname|dsname|fullxres|fullyres|&value=3|\"vcodec\"|\"folder\"" | sed -e "s|10.10.1.45/JSON?user=userName&pass=P@ssw0rd&request=controldevicebyvalue&ref=||g" | sed -e "s|&value=3||g" > ${host}.raw
	echo ======== Raw =========
	cat ${host}.raw
	echo ======== Raw =========

	cut -f2-20 -d'=' ${host}.raw | sed -e "s|
||g" | sed -e "s|cameraId=1538607685:1\"|cameraId=1538607685:1\"\ndword:00000280\ndword:000001e0|g" | sed -e "s|dword:||g" > ${host}.tmp
	echo ======== Tmp =========
	cat ${host}.tmp
	echo ======== Tmp =========

	fmt="$host,%s,%s\\n"
	awk -v fmt="$host,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,\\n" '{getline b;getline c;getline d;getline e;getline f;getline g;getline h;getline i;getline j;getline k;getline l;getline m;getline n;getline o;getline p;getline q;getline r;printf(fmt,c,d,e,f,g,h,$0,p,q,r,b)}' ${host}.tmp 2>&1 >> cams.csv
	echo ======== Cams =========
done
cat cams.csv

read ans

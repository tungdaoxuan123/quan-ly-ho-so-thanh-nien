@echo off
rem Opens port 3000 for phones on the Wi-Fi, on every network profile: a Wi-Fi without internet is usually
rem "Unidentified network" (Public), which the usual "Allow access" prompt for Private networks does not cover.
echo Mo cong 3000 tren tuong lua Windows de dien thoai trong wifi vao duoc ung dung.
netsh advfirewall firewall delete rule name="Quan Ly Ho So" >nul 2>&1
netsh advfirewall firewall add rule name="Quan Ly Ho So" dir=in action=allow protocol=TCP localport=3000 profile=any
if errorlevel 1 (
  echo.
  echo KHONG THANH CONG: bam chuot phai vao tep nay va chon "Run as administrator".
) else (
  echo.
  echo Da mo cong 3000. Chi can lam mot lan tren may nay.
)
pause

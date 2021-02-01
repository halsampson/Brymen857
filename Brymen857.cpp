// Brymen BM857 multimeter optical interface
 
#include "windows.h"
#include <stdio.h>
#include "string.h"
#include "setupapi.h"
#pragma comment(lib, "setupAPI.lib")

const int Brymen857Baud = 1000000 / 128;  // 128us per bit

const char* lastActiveComPort() { // to default to last USB serial adapter plugged in
  static char comPortName[8] = "none";
  HDEVINFO hDevInfo = SetupDiGetClassDevs(&GUID_CLASS_COMPORT, NULL, NULL, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
  FILETIME latest = { 0 };

  for (int index = 0; ; index++) {
    SP_DEVINFO_DATA DeviceInfoData;
    DeviceInfoData.cbSize = sizeof(DeviceInfoData);
    if (!SetupDiEnumDeviceInfo(hDevInfo, index, &DeviceInfoData)) break;

    char hardwareID[256];
    SetupDiGetDeviceRegistryProperty(hDevInfo, &DeviceInfoData, SPDRP_HARDWAREID, NULL, (BYTE*)hardwareID, sizeof(hardwareID), NULL);

    // truncate to make registry key for common USB CDC devices:
    char* truncate;
    if ((truncate = strstr(hardwareID, "&REV"))) *truncate = 0;
    if ((truncate = strstr(hardwareID, "\\COMPORT"))) *truncate = 0;  // FTDI 

    char devKeyName[256] = "System\\CurrentControlSet\\Enum\\";
    strcat_s(devKeyName, sizeof(devKeyName), hardwareID);
    HKEY devKey = 0;
    LSTATUS res = RegOpenKeyEx(HKEY_LOCAL_MACHINE, devKeyName, 0, KEY_READ, &devKey);
    if (devKey) {
      DWORD idx = 0;
      while (1) {
        char serNum[64]; DWORD len = sizeof(serNum);
        FILETIME lastWritten = { 0 }; // 100 ns
        if (RegEnumKeyEx(devKey, idx++, serNum, &len, NULL, NULL, NULL, &lastWritten)) break;
        if (CompareFileTime(&lastWritten, &latest) > 0) { // latest device connected
          latest = lastWritten;
          if (strstr(devKeyName, "FTDIBUS")) strcat_s(serNum, sizeof(serNum), "\\0000"); // TODO: enumerate FTDI?
          strcat_s(serNum, sizeof(serNum), "\\Device Parameters");
          len = sizeof(comPortName);
          RegGetValue(devKey, serNum, "PortName", RRF_RT_REG_SZ, NULL, comPortName, &len);
          // SetupDiGetDeviceRegistryProperty(hDevInfo, &DeviceInfoData, SPDRP_FRIENDLYNAME, NULL, (BYTE*)devName, sizeof(devName), NULL);        
        }
      }
      RegCloseKey(devKey);
    }
    else if (strstr(hardwareID, "USB")) printf("%s not found", devKeyName); // low numbered non-removable ports ignored
  }

  SetupDiDestroyDeviceInfoList(hDevInfo);
  return comPortName;

  // see also HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM
  //   can check if disconnect removes from list
  //   can't tell which is last enumerated
}

HANDLE hCom;
bool openSerial(const char* portName) {
  char portDev[16] = "\\\\.\\";
  strcat_s(portDev, sizeof(portDev), portName);
  hCom = CreateFileA(portDev, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, NULL, NULL);
  if (hCom == INVALID_HANDLE_VALUE) return false;

  DCB dcb = { 0 };
  dcb.DCBlength = sizeof(DCB);
  dcb.BaudRate = Brymen857Baud;
  dcb.ByteSize = 8;
  dcb.fBinary = TRUE;
  if (!SetCommState(hCom, &dcb)) return false;
  if (!SetupComm(hCom, 16, 64)) return false;

  COMMTIMEOUTS timeouts = { 0 };  // in ms
  timeouts.ReadTotalTimeoutConstant = 1000 + 265;  // for Hz + more for large capacitance (rare)
  timeouts.ReadIntervalTimeout = 64;  // bulk USB 64 byte partial buffer timeout
  if (!SetCommTimeouts(hCom, &timeouts)) return false;

  return true;
}

const int RawLen = 35;

typedef struct {
  BYTE start : 2;  // always 1
  BYTE data  : 4;
  BYTE stop  : 2;  // always 3
} rawData;

rawData raw[RawLen];

// TODO: unknown LCD bits

struct {
  union {
    BYTE switchPos;
    struct {
      BYTE swPosLSN : 4;
      BYTE swPosMSN : 4;
    };
  };

  BYTE unk00 : 1;
  BYTE milli : 1;
  BYTE micro : 1;
  BYTE Volts : 1;
  BYTE Hz    : 1;
  BYTE Ohms  : 1;
  BYTE kilo  : 1;
  BYTE Mega  : 1;

  BYTE Auto : 1;
  BYTE Hold : 1;
  BYTE DC : 1;
  BYTE AC : 1;
  BYTE unk45 : 2;
  BYTE nano : 1;
  BYTE unk7 : 1;

  struct {
    BYTE segments : 7;
    BYTE decimalOrMinus : 1;
  } lcdDigits[6];

  BYTE bar0to5 : 6;
  BYTE beep : 1;
  BYTE unk1 : 1;

  BYTE bargraph[5]; // 6 to 45

  BYTE unk2 : 1;
  BYTE Min  : 1;
  BYTE unk3 : 1;
  BYTE percent : 1;
  BYTE bar46to49 : 4;

  BYTE unk4 : 4;
  BYTE minusMaxMin : 1;
  BYTE unk5 : 1;
  BYTE Max : 1;
  BYTE Rec : 1;

  BYTE unk6 : 7;
  BYTE delta : 1;
} packed;


double decodeRaw() {
  for (int b = 0; b < RawLen; b++)
    if (raw[b].start != 1 || raw[b].stop != 3) printf("Err %d ", b);  // only 4 bits vary

  BYTE* p = (BYTE*)&packed;
  rawData* pRaw = raw + 1;
  do {
    *p = (*pRaw++).data << 4;
    *p++ |= (*pRaw++).data;
  } while (pRaw < raw + sizeof(raw));

  const char LCDchars[] = "0123456789ELnr-";
  const char Segments[] = {
     0x7D,    5, 0x5B, 0x1F, 0x27, // 0..4
     0x3E, 0x7E, 0x15, 0x7F, 0x3F, // 5..9
     0x7A, 0x68, 0x46, 0x42, 2,    // ELnr-
     0
  };

  char numStr[8];
  char* pNum = numStr;
  for (int digit = 0; digit < sizeof(packed.lcdDigits); digit++) {
    if (packed.lcdDigits[digit].decimalOrMinus) *pNum++ = digit ? '.' : '-';
    if (packed.lcdDigits[digit].segments) { // not blank
      const char* lcdChar = strchr(Segments, (char)(packed.lcdDigits[digit].segments));
      if (lcdChar) *pNum++ = LCDchars[lcdChar - Segments];
      // else printf("%X ", packed.digits[digit] & 0x7F); // error
    }
  }
  *pNum = 0;
  double reading = atof(numStr);

  const char* units[] = { "V", "V", "V", "Hz", "V", "Ohm", "F", "A", "A", "dBm", "%%" };
  packed.switchPos = ~packed.switchPos;  // single bit
  BYTE switchPos = 0;
  if (packed.switchPos) {
    const int log2[9] = { 0, 1, 2, 2, 3, 3, 3, 3, 4 };
    if (packed.swPosMSN)
         switchPos = 5 - log2[packed.swPosMSN];
    else switchPos = 4 + log2[packed.swPosLSN];
  }

  const char* range = "";
  if (switchPos == 0 && !packed.Volts) switchPos = 9; // dBm
  else if (packed.percent) switchPos = 10; // %
  else {
    if (packed.nano)  { range = "n"; reading *= 1E-9; }
    if (packed.micro) { range = "u"; reading *= 1E-6; }
    if (packed.milli) { range = "m"; reading *= 1E-3; }
    if (packed.kilo)  { range = "k"; reading *= 1E3; }
    if (packed.Mega)  { range = "M"; reading *= 1E6; }
  }

  const char* acdc = "";
  if (packed.DC) acdc = "DC";
  if (packed.AC) acdc = "AC";
  if (packed.AC && packed.DC) acdc = "AC+DC";

  const char* modifier = "";
  if (packed.Min) modifier = " Min";
  if (packed.Max) modifier = " Max";
  if (packed.Rec) modifier = " Rec";
  if (packed.delta) modifier = " delta";

  // more LCD bits:
  // printf("  %X %X %X %X", packed.range[1], packed.percent, packed.record, packed.delta); 

  printf("%s %s%s%s %s\n", numStr, range, units[switchPos], acdc, modifier);

  return reading;
}


int main(int argc, char** argv) {
  const char* comPort = argc > 1 ? argv[1] : lastActiveComPort();
  if (openSerial(comPort)) {
    printf("Connected to %s\n", comPort);
  } else {
    printf("(Re)connect USB serial adapter or use: Brymen857 COMnn\n");
    return -2;
  }

  while (1) {
    bool ret = SetCommBreak(hCom); // TxD low -> IRED on
    if (!ret) printf("SetCommBreak not supported\n");
    Sleep(10);
    ret = ClearCommBreak(hCom); // Txd high -> IRED off
      // NOTE: can leave Tx IRED on constantly

    DWORD bytesRead;
    if (!ReadFile(hCom, raw, RawLen, &bytesRead, NULL)) break;
    if (bytesRead == RawLen) decodeRaw();
 
    Sleep(20);
  }

  if (hCom) CloseHandle(hCom);
}
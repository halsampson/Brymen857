// Brymen BM857 multimeter optical interface
 
#include "windows.h"
#include <stdio.h>
#include <conio.h>
#include <math.h>
#include <string.h>
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
  //   to check disconnect removed from list
  //   can't tell which is last enumerated
}

HANDLE openSerial(const char* portName) {
  char portDev[16] = "\\\\.\\";
  strcat_s(portDev, sizeof(portDev), portName);
  HANDLE hCom = CreateFileA(portDev, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING, NULL, NULL);
  if (hCom == INVALID_HANDLE_VALUE) return hCom;

  DCB dcb = { 0 };
  dcb.DCBlength = sizeof(DCB);
  dcb.BaudRate = Brymen857Baud;
  dcb.ByteSize = 8;
  dcb.fBinary = TRUE;
  if (!SetCommState(hCom, &dcb)) return 0;
  if (!SetupComm(hCom, 16, 64)) return 0;

  COMMTIMEOUTS timeouts = { 0 };  // in ms
  timeouts.ReadTotalTimeoutConstant = 3000 + 265;  // for Hz + more for large capacitance (rare TODO)
  timeouts.ReadIntervalTimeout = 64;  // bulk USB 64 byte partial buffer timeout
  if (!SetCommTimeouts(hCom, &timeouts)) return 0;

  return hCom;
}

const int RawLen = 35;

typedef struct {
  BYTE start : 2;  // always 1
  BYTE data  : 4;
  BYTE stop  : 2;  // always 3
} rawData;

rawData raw[RawLen];


// TODO: fill in unknown LCD bits

struct {
  union {
    BYTE switchPos;
    struct {
      BYTE swPosLSN : 4;
      BYTE swPosMSN : 4;
    };
  };

  BYTE dBm   : 1;
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
  BYTE unk7 : 2; // V; A?
  BYTE nano : 1;
  BYTE unk8 : 1; // FS?

  struct {
    BYTE segments  : 7;
    BYTE dpOrMinus : 1;
  } lcdDigits[6];

  BYTE bar0to5 : 6;
  BYTE beep : 1;
  BYTE unk1 : 1;  // LowBatt?

  BYTE bargraph[5]; // 6 to 45

  BYTE unk2 : 1; 
  BYTE Min  : 1;
  BYTE unk3 : 1;  // Avg ?
  BYTE percent : 1;
  BYTE bar46to49 : 4;

  BYTE unk4 : 4;
  BYTE PeaktoPeak : 1; // ??
  BYTE unk5 : 1; 
  BYTE Max : 1;
  BYTE Delta : 1;
  // Crest ?

  BYTE unk6 : 8;
} packed;


bool packRaw(void) {
  for (int b = 0; b < RawLen; b++)
    if (raw[b].start != 1 || raw[b].stop != 3) {
      printf("Err %d ", b);  // only 4 bits vary
      return false;
    }

  BYTE* p = (BYTE*)&packed;
  rawData* pRaw = raw + 1;
  do { // pack raw nibbles into packed bytes
    *p = (*pRaw++).data << 4;
    *p++ |= (*pRaw++).data;
  } while (pRaw < raw + sizeof(raw));

  packed.switchPos = ~packed.switchPos;  // single bit sent inverted
  return true;
}


char numStr[8];
const char* range;

double getLcdValue(void) {
  const char LCDchars[] = "0123456789ELnr-";
  const char Segments[] = { // 7 segment bits
     0x7D,    5, 0x5B, 0x1F, 0x27, // 0..4
     0x3E, 0x7E, 0x15, 0x7F, 0x3F, // 5..9
     0x7A, 0x68, 0x46, 0x42, 2,    // ELnr-
     0
  };

  char* pNum = numStr;
  for (int digit = 0; digit < sizeof(packed.lcdDigits); digit++) {
    if (packed.lcdDigits[digit].dpOrMinus) *pNum++ = digit ? '.' : '-';
    if (packed.lcdDigits[digit].segments) { // not blank
      const char* lcdChar = strchr(Segments, (char)(packed.lcdDigits[digit].segments));
      if (lcdChar) *pNum++ = LCDchars[lcdChar - Segments];
    }
  }
  *pNum = 0;
  double value = atof(numStr);

  range = "";
  if (!packed.dBm && !packed.percent) {
    if (packed.nano)  { range = "n"; value *= 1E-9; }
    if (packed.micro) { range = "u"; value *= 1E-6; }
    if (packed.milli) { range = "m"; value *= 1E-3; }
    if (packed.kilo)  { range = "k"; value *= 1E3; }
    if (packed.Mega)  { range = "M"; value *= 1E6; }
  }
  return value;
}


const char* units;
const char* acdc;
const char* modifier;

void getUnits(void) {  // sets 3 strings above
  const char* unitStr[] = { "V", "V", "V", "Hz", "V", "Ohm", "F", "A", "A", "dBm", "%" };
  BYTE switchPos = 0;
  if (packed.switchPos) {
    const int log2[9] = { 0, 1, 2, 2, 3, 3, 3, 3, 4 }; // bit position
    if (packed.swPosMSN)
      switchPos = 5 - log2[packed.swPosMSN];
    else switchPos = 4 + log2[packed.swPosLSN];
  }
  if (packed.dBm) switchPos = 9; // dBm  
  else if (packed.percent) switchPos = 10; // %

  units = unitStr[switchPos];

  acdc = "";
  if (packed.AC && packed.DC) {
    acdc = " AC+DC";
  } else {
    if (packed.DC) acdc = "DC";
    if (packed.AC) acdc = "AC";
  }

  modifier = "";
  if (packed.Min) modifier = "Min";
  if (packed.Max) modifier = "Max";
  if (packed.PeaktoPeak) modifier = "p-p"; // ?not seen
  if (packed.Delta) modifier = "Delta"; 
  // if (packed.Rec) modifier = "Rec";
  
  // unknown LCD bits:
  // printf("%X %X%X%X%X %X%X%X%X ", raw[0].data, packed.unk1, packed.unk2, packed.unk3, packed.unk4, packed.unk5, packed.unk6, packed.unk7, packed.unk8);
}

const double MinErrVal = 1E9;

double decodeRaw(bool doUnits = true) {
  if (!packRaw()) return 9E9;

  double reading = getLcdValue();
  if (doUnits) getUnits();
  return reading;
}

double getReading(HANDLE hCom) {
  bool ret = SetCommBreak(hCom); // TxD low -> IRED on
  if (!ret) printf("Break not supported\n");
  // Sleep(10);
  // ret = ClearCommBreak(hCom); // Txd high -> IRED off
  // NOTE: can leave Tx IRED on constantly

  DWORD bytesRead;
  if (!ReadFile(hCom, raw, RawLen, &bytesRead, NULL)) return 0;
  if (bytesRead != RawLen) return 9E9;  // TODO: retry
  
  return decodeRaw();
}

HANDLE hBrymen, hPSU;

void sendPSUCmd(const char* cmd) {
  WriteFile(hPSU, cmd, (DWORD)strlen(cmd), NULL, NULL);
}

char response[64];

char* getResponse(const char* cmd) {
  sendPSUCmd(cmd);
  DWORD bytesRcvd = 0;
  if (ReadFile(hPSU, response, sizeof(response) - 1, &bytesRcvd, NULL))
    response[bytesRcvd] = 0;
  return response;
}

int getValue(const char* cmd) {
  return atoi(getResponse(cmd));
}

char cmd[32];
int settle_ms = 8000;

// set to range of interest:
const double VcalLo = 3;
const double VcalHi = 5;  // 50000/500000 counts

const int NAVG = 1024 * 2 + 1;  // 400 ms? sampling beats with 60 Hz
double val[NAVG];

// beware synchronizing with noise???
double avgReading(HANDLE hMeter) { // median
  sendPSUCmd(cmd);
  Sleep(settle_ms);

  while (1) {    
    for (int n = 0; n < NAVG; n++) {
      double newval;
      do newval = getReading(hMeter); while (newval >= MinErrVal);
      if (0) { // sort
        int newpos = n;
        for (int i = 0; i < n; i++)
          if (newval < val[i]) {   // keep sorted
            newpos = i;
            for (int j = n; j > i; j--)
              val[j] = val[j - 1];
            break;
          }
        val[newpos] = newval;
      } else val[n] = newval;
    }

    double p2p = 1E6 * fabs(val[NAVG - 1] - val[0]);
    printf(" %.5f %.5f", val[0], val[NAVG - 1]);
    if (p2p > 2000) {
      printf("!");
      Sleep(settle_ms);
    } else 
      return val[NAVG / 2]; // median
  }
}

const int DAC_SCALE = 4096;
int dacLo, dacHi;
double readLoV, readHiV;
int maxVout[5], offset;
int ref;

void readLoHiV() {
  sprintf_s(cmd, sizeof(cmd), "+%dD", dacLo = int(DAC_SCALE * (VcalLo * 1E6 + offset) / maxVout[ref] + 0.5));
  readLoV = avgReading(hBrymen);
  printf(" %.5f", readLoV);

  sprintf_s(cmd, sizeof(cmd), "+%dD", dacHi = int(DAC_SCALE * (VcalHi * 1E6 + offset) / maxVout[ref] + 0.5));
  readHiV = avgReading(hBrymen);
  printf(" %.5f", readHiV);
}

#pragma warning( disable : 26451)
#pragma warning( disable : 6387)

int board;

void calibrateVref(bool adjust = true) { // and offset 
  for (ref = 0; ref < 1; ref += 4) {  // 1.25 (and Vdd)
    Sleep(100);
    sprintf_s(cmd, sizeof(cmd), "+%d ", ref); // set param0 = REFSEL
    sendPSUCmd(cmd);
    Sleep(100);

    // get calibration
    offset = getValue("o");
    int prevmaxVout = maxVout[ref] = getValue("b");

    double temperature = getValue("t") / 100.;
    
    readLoHiV();
    if (adjust) {
      maxVout[ref] = int(1E6 * (readHiV - readLoV) * DAC_SCALE / (dacHi - dacLo) + 0.5); // uV
      offset = -int(1E6 * (readLoV * dacHi - readHiV * dacLo) / (dacHi - dacLo) + 0.5); // uV at output
      // char cmd[32];
      // sprintf_s(cmd, sizeof(cmd), "+%dO+%dB", offset, maxVout[ref]);
      // sendPSUCmd(cmd);

      char calStr[64];
      int len = sprintf_s(calStr, sizeof(calStr), "%d, %d, %.2f, %6d, %d, %d\n",
        board, ref, temperature, offset, maxVout[ref], 1000000 * (maxVout[ref] - prevmaxVout) / maxVout[ref]); // PPM change
      printf("  %s", calStr);

      FILE* fCalib;
      if (!fopen_s(&fCalib, "calibt.csv", "a+t")) {
        fwrite(calStr, 1, len, fCalib);
        fclose(fCalib);
      }
    }
  }
#if 0
  for (ref = 0; ref < 5; ref++)
    printf("%d, ", maxVout[ref]);
  printf("\n\n");
#endif
}

void calibratePowerSupply(void) {
  bool adjust = true;
  while (1) {
    do calibrateVref(adjust);
    while (!_kbhit());
    adjust = _getch() == 'a';
  }

  for (double volts = 0.5; volts <= 36; volts += max(0.02, volts / 30)) {
    char setVolts[16];
    sprintf_s(setVolts, sizeof(setVolts), "%.3fV", volts);
    WriteFile(hPSU, setVolts, (DWORD)strlen(setVolts), NULL, NULL);
    Sleep(3000);
    double reading = getReading(hBrymen);
    printf("%.3f, %.5f, %.5f\n", volts, volts - reading, reading);
    // to CSV file also
  }
}

void powerSupplyTest(bool brymenOK) {
  sendPSUCmd("0I"); // Interactive off
  sendPSUCmd("0T");  // heater noise off
  printf("%s\n", getResponse("i")); // identify
  char* beforeNum = strrchr(response, ' ');
  if (beforeNum) {
    board = atoi(beforeNum + 1);
    if (brymenOK)
      calibratePowerSupply();
    // else chipTempTest();
  }
  CloseHandle(hPSU); 
}

signed char ofs[100000];

int main(int argc, char** argv) {
  const char* comPort = argc > 1 ? argv[1] : lastActiveComPort();
  if ((hBrymen = openSerial(comPort)) <= 0) {
    printf("(Re)connect USB serial adapter or use:  Brymen857 COMnn\n");
    return -2;
  }
  bool brymenOK = getReading(hBrymen) < MinErrVal;
  if (brymenOK)
    printf("Brymen connected on %s\n", comPort);

  int maxR = 0, minR = 0;

  for (int i = 0; i < sizeof(ofs); i++) {
    ofs[i] = (int)(getReading(hBrymen) * 1E5);
    if (ofs[i] > maxR) {
      maxR = ofs[i];
      printf("%d ", maxR);
    }
    if (ofs[i] < minR) {
      minR = ofs[i];
      printf("%d ", minR);
    }
  }

  while (1);


  if ((hPSU = openSerial(argc > 2 ? argv[2] : lastActiveComPort())) != INVALID_HANDLE_VALUE
  || (hPSU = openSerial("COM7")) != INVALID_HANDLE_VALUE
  || (hPSU = openSerial("COM8")) != INVALID_HANDLE_VALUE)
    powerSupplyTest(brymenOK);

  if (brymenOK) while (1) {
    double reading = getReading(hBrymen);
    if (reading < MinErrVal)
      printf("%s %s%s %s %s\n", numStr, range, units, acdc, modifier);    
  } 
  
  CloseHandle(hBrymen);
}
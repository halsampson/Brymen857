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
  timeouts.ReadTotalTimeoutConstant = 1000 + 265;  // for Hz + more for large capacitance (rare TODO)
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

const double MaxErrVal = -8E88;

double decodeRaw(bool doUnits = true) {
  if (!packRaw()) return -8E88;
  double reading = getLcdValue();
  if (doUnits) getUnits();
  return reading;
}

double getReading(HANDLE hCom) {
  bool ret = SetCommBreak(hCom); // TxD low -> IRED on
  if (!ret) printf("Break not supported\n");
  Sleep(1);
  ret = ClearCommBreak(hCom); // Txd high -> IRED off
    // NOTE: can leave Tx IRED on constantly

  DWORD bytesRead;
  if (!ReadFile(hCom, raw, RawLen, &bytesRead, NULL)) return 0;
  if (bytesRead != RawLen) return -9E99;

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
  if (ReadFile(hPSU, response, sizeof(response), &bytesRcvd, NULL))
    response[bytesRcvd] = 0;
  return response;
}

int getValue(const char* cmd) {
  return atoi(getResponse(cmd));
}

// set to range of interest:
int VcalLo = 1;
int VcalHi = 10;

void calibrateResistor(void) {  // stable: only rarely
  // get calibration
  int offset = getValue("o");
  int R10K = getValue("r");

  int settle_ms = 5000;
  while (1) {
    printf("%d %d\n", offset, R10K);
    if (settle_ms > 10000) break;

    char cmd[32];
    sprintf_s(cmd, sizeof(cmd), "%dO%dR%dV", offset, R10K, VcalLo);
    sendPSUCmd(cmd);
    Sleep(settle_ms);
    double readLoV = getReading(hBrymen);

    sprintf_s(cmd, sizeof(cmd), "%dV", VcalHi);
    sendPSUCmd(cmd);
    Sleep(settle_ms);
    double readHiV = getReading(hBrymen);

    R10K = int((R10K + 1000) * (readHiV - readLoV) / (VcalHi - VcalLo) - 1000 + 0.5);
    const double OffsetVoltsPerCount = 11 * 1.25 / 16384;
    offset += int((VcalLo - readLoV) / OffsetVoltsPerCount + 0.5); // added to target ADC

    settle_ms += settle_ms / 4;
  }
}

int board;

void calibrateVref(void) { // and offset 
  double temperature = getValue("j") / 100.;
  printf("%.2fC\n", temperature);

  // get calibration
  int offset = getValue("o");
  int ref2p50V = getValue("b");

  int settle_ms = 5000;
  while (1) {
    printf("%d %d\n", offset, ref2p50V);
    if (settle_ms > 10000) break;

    char cmd[32];
    sprintf_s(cmd, sizeof(cmd), "%dO%dB", offset, ref2p50V);
    sendPSUCmd(cmd);
    Sleep(2000);  // wait for calibration averaging

    sprintf_s(cmd, sizeof(cmd), "%dV", VcalLo);
    sendPSUCmd(cmd);
    Sleep(settle_ms);
    double readLoV = getReading(hBrymen);

    sprintf_s(cmd, sizeof(cmd), "%dV", VcalHi);
    sendPSUCmd(cmd);
    Sleep(settle_ms);
    double readHiV = getReading(hBrymen);

    ref2p50V = int(ref2p50V * (readHiV - readLoV) / (VcalHi - VcalLo) + 0.5);
    const double OffsetVoltsPerCount = 11 * 1.25 / 16384;
    offset += int((VcalLo - readLoV) / OffsetVoltsPerCount + 0.5); // added to target ADC

    settle_ms += settle_ms / 4;
  }

  FILE* fCalib;
  if (!fopen_s(&fCalib, "calib.csv", "a+t")) {
    char calStr[64];
    int len = sprintf_s(calStr, sizeof(calStr), "%d, %.2f, %d, %d\n", board, temperature, offset, ref2p50V);
    fwrite(calStr, 1, len, fCalib);
    fclose(fCalib);
  }
}

void calibratePowerSupply(char* psuComPort) {
  if ((hPSU = openSerial(psuComPort)) <= 0) return;

  getResponse("0E"); // Echo off
  printf("%s\n", getResponse("i")); // identify
  board = atoi(strrchr(response, ' ') + 1);

  calibrateVref();

  for (double volts = 0.5; volts <= 36; volts += max(0.02, volts / 30)) {
    char setVolts[16];
    sprintf_s(setVolts, sizeof(setVolts), "%.3fV", volts);
    WriteFile(hPSU, setVolts, (DWORD)strlen(setVolts), NULL, NULL);
    Sleep(3000);
    double reading = getReading(hBrymen);
    printf("%.3f, %.5f, %.5f\n", volts, volts - reading, reading);
    // to CSV file also
  }
  CloseHandle(hPSU);  
}

int main(int argc, char** argv) {
  const char* comPort = argc > 1 ? argv[1] : lastActiveComPort();
  if ((hBrymen = openSerial(comPort)) <= 0) {
    printf("(Re)connect USB serial adapter or use:  Brymen857 COMnn\n");
    return -2;
  }
  printf("Connected to %s\n", comPort);

  if (argc <= 2) {
    while (1) {
      double reading = getReading(hBrymen);
      if (reading > MaxErrVal)
        printf("%s %s%s %s %s\n", numStr, range, units, acdc, modifier);
    }
  } else calibratePowerSupply(argv[2]);

  CloseHandle(hBrymen);
}
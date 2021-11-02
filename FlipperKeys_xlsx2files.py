#!/usr/bin/python3
#
# Convert xlsx-file of Keys to Flipper files
# The script handles keys: Dallas, Cyfral, Metakom, EM-Marin, Indala, HID Prox
# 
# Author: Serega Barsukov senbarsukov@gmail.com
#
# Head of xlsx:
# =============
# № | Key Type | UID | Comment | <empty> | Extra flags
#
# Flipper iButton File example:
# =============================
# Filetype: Flipper iButton key
# Version: 1
# # Key type can be Cyfral, Dallas or Metakom
# Key type: Dallas
# # Data size for Cyfral is 2, for Metakom is 4, for Dallas is 8
# Data: 01 11 12 22 23 33 34 40
#
# Flipper RFID File eaxample:
# ===========================
# Filetype: Flipper RFID key
# Version: 1
# # Key type can be EM4100, H10301 or I40134
# Key type: EM4100
# # Data size for EM4100 is 5, for H10301 is 3, for I40134 is 3
# Data: 11 11 11 11 11

import os
import re
import openpyxl
import string
import transliterate
from pathlib import Path

# Inputs
FILE_NAME='RFID_and_iButton_keys.xlsx'
PATH=Path().resolve()

# Class
class Keys_Xlsx2FlipperFiles:
    def __init__(self, pa_sFileName):
        self.FILE_NAME = pa_sFileName
        # Constants
        self.__NAME_CHAR_LIMIT=21
        self.__KEY_TYPE_DALLAS=("dallas",)
        self.__KEY_TYPE_CYFRAL=("cyfral",)
        self.__KEY_TYPE_METAKOM=("metakom",)
        self.__KEY_TYPE_EMMARIN=("em marin", "em-marin")
        self.__KEY_TYPE_INDALA=("indala",)
        self.__KEY_TYPE_HIDPROX=("hid prox", "hid-prox")
        self.__FLIPP_NAME_DALLAS="Dallas"
        self.__FLIPP_NAME_CYFRAL="Cyfral"
        self.__FLIPP_NAME_METAKOM="Metakom"
        self.__FLIPP_NAME_EMMARIN="EM4100"
        self.__FLIPP_NAME_INDALA="I40134"
        self.__FLIPP_NAME_HIDPROX="H10301"
        self.__KEY_TYPE_IBUTTON={self.__FLIPP_NAME_DALLAS: self.__KEY_TYPE_DALLAS, 
                                 self.__FLIPP_NAME_CYFRAL: self.__KEY_TYPE_CYFRAL, 
                                 self.__FLIPP_NAME_METAKOM: self.__KEY_TYPE_METAKOM}
        self.__KEY_TYPE_125RFID={self.__FLIPP_NAME_EMMARIN: self.__KEY_TYPE_EMMARIN, 
                                 self.__FLIPP_NAME_INDALA: self.__KEY_TYPE_INDALA,
                                 self.__FLIPP_NAME_HIDPROX: self.__KEY_TYPE_HIDPROX}
        self.__KEY_DATA_SIZE_IBUTTON={self.__FLIPP_NAME_DALLAS: 8,
                                      self.__FLIPP_NAME_CYFRAL: 2,
                                      self.__FLIPP_NAME_METAKOM: 4}
        self.__KEY_DATA_SIZE_RFID={self.__FLIPP_NAME_EMMARIN: 5,
                                   self.__FLIPP_NAME_INDALA: 3,
                                   self.__FLIPP_NAME_HIDPROX: 3}

        # Read the active sheet:
        sFile = Path(pa_sFileName)
        oWb = openpyxl.load_workbook(sFile) 
        self.__sheet = oWb.active

        self.__dirIButton = os.path.join(PATH, 'iButton')
        if not os.path.exists(self.__dirIButton):
            os.mkdir(self.__dirIButton)

        self.__dirRFID = os.path.join(PATH, 'RFID')
        if not os.path.exists(self.__dirRFID):
            os.mkdir(self.__dirRFID)

        self.__listOfFilesIButton={}
        self.__listOfFilesRFID={}

    def __getStartRow(self) -> int:
      '''
      Get first Row to handle. 
      Searching for '1' in '№' column
      '''
      nCnt = 1
      while(self.__sheet['A'+str(nCnt)].value != 1):
        nCnt += 1
      return nCnt

    
    def __isKnownKey(self, pa_sKeyName: str, pa_oKeyTypeDict: dict):
        sKeyName = pa_sKeyName.lower()
        for flippKeyName, keyNameTuple in pa_oKeyTypeDict.items():
            if len(keyNameTuple) == 1:
                if str(keyNameTuple[0]) in sKeyName:
                   return flippKeyName 
            else:
                for severalNames in keyNameTuple:
                    if str(severalNames) in sKeyName:
                        return flippKeyName
        return None
    
    def __createFile_iButton(self, pa_sData: str, pa_sKeyType: str, pa_sFileName: str):
        sFilePath = os.path.join(self.__dirIButton, pa_sFileName)
        if os.path.isfile(sFilePath):
            os.remove(sFilePath)
        oFile = open(sFilePath, 'x')
        oFile.write("Filetype: Flipper iButton key\n")
        oFile.write("Version: 1\n")
        oFile.write("# Key type can be Cyfral, Dallas or Metakom\n")
        oFile.write(f"Key type: {pa_sKeyType}\n")
        oFile.write("# Data size for Cyfral is 2, for Metakom is 4, for Dallas is 8\n")
        oFile.write(f"Data: {pa_sData}\n")
        oFile.close

    def __createFile_RFID(self, pa_sData: str, pa_sKeyType: str, pa_sFileName: str):
        sFilePath = os.path.join(self.__dirRFID, pa_sFileName)
        if os.path.isfile(sFilePath):
            os.remove(sFilePath)
        oFile = open(sFilePath, 'x')
        oFile.write("Filetype: Flipper RFID key\n")
        oFile.write("Version: 1\n")
        oFile.write("# Key type can be EM4100, H10301 or I40134\n")
        oFile.write(f"Key type: {pa_sKeyType}\n")
        oFile.write("# Data size for EM4100 is 5, for H10301 is 3, for I40134 is 3\n")
        oFile.write(f"Data: {pa_sData}\n")
        oFile.close

    def __renameExistingFile(self, pa_sFileName: str, pa_lNameDict: dict) -> str:
        nNameResizeVal = self.__NAME_CHAR_LIMIT - 5
        if len(pa_sFileName) > nNameResizeVal:
            pa_sFileName = pa_sFileName[:nNmaeResizeVal]
        return pa_sFileName + "_{:03d}_".format(pa_lNameDict[pa_sFileName])

    def _fileNameHandling(self, pa_sFileName: str, pa_lNameDict: dict) -> str:
        pa_sFileName = transliterate.translit(pa_sFileName, 'ru', reversed=True)
        tmpList = pa_sFileName.split()
        pa_sFileName = '_'.join(tmpList)
        if len(pa_sFileName) > self.__NAME_CHAR_LIMIT:
            pa_sFileName = pa_sFileName[:self.__NAME_CHAR_LIMIT]
        pa_sFileName = re.sub('\W', 'X', pa_sFileName)
        
        if pa_sFileName in pa_lNameDict:
            pa_lNameDict[pa_sFileName] += 1
            pa_sFileName = self.__renameExistingFile(pa_sFileName, pa_lNameDict)
            pa_lNameDict[pa_sFileName] = 1
        else:
            pa_lNameDict[pa_sFileName] = 1
        return pa_sFileName

    def __getOnlyHexData(self, lData: list) -> list:
        err = 0
        lRmFromList = []
        for d in lData:
            for i, char in enumerate(d):
                if not char in string.hexdigits or i > 1:
                    lRmFromList.append(d)
                    break

        for l in lRmFromList:
            lData.remove(l)
        return lData

    def _checkKeyData(self, sData: str, sDataType: str, lKeyDataSize: dict):
        nKeyDataSize = lKeyDataSize[sDataType]

        tmpData = str(sData).split()
        tmpData = self.__getOnlyHexData(tmpData)
        nInputDataSize = len(tmpData)
        
        tmpData = (" ".join(tmpData))

        if nInputDataSize == nKeyDataSize:
            return tmpData
        
        # not correct data size
        return None


    def convert(self) -> int:
        nErr = 0
        nStartRow = self.__getStartRow()
        while True:
            sKeyName = self.__sheet['B'+str(nStartRow)].value
            if type(sKeyName) == type(None):
                ## End of File or Space
                return nErr
            
            sIButtonKeyType = (self.__isKnownKey(sKeyName, self.__KEY_TYPE_IBUTTON))
            if sIButtonKeyType == None:
                sRFIDKeyType = (self.__isKnownKey(sKeyName, self.__KEY_TYPE_125RFID))
                if sRFIDKeyType:
                    ## RFID
                    sKeyData = self._checkKeyData(self.__sheet['C'+str(nStartRow)].value,
                                                   sRFIDKeyType,
                                                   self.__KEY_DATA_SIZE_RFID)
                    if sKeyData == None:
                        print("ERROR::RFID::Key-Data is not equal to Key-Type")
                        nErr += 1

                    sFileName = self._fileNameHandling(self.__sheet['D'+str(nStartRow)].value,
                                                       self.__listOfFilesIButton)
                    self.__createFile_RFID(sKeyData,
                                           sRFIDKeyType,
                                           sFileName+".rfid")
            else:
                ## iButton
                sKeyData = self._checkKeyData(self.__sheet['C'+str(nStartRow)].value,
                                               sIButtonKeyType,
                                               self.__KEY_DATA_SIZE_IBUTTON)
                if sKeyData == None:
                    print("ERROR::IBUTTON::Key-Data is not equal to Key-Type")
                    nErr += 1

                sFileName = self._fileNameHandling(self.__sheet['D'+str(nStartRow)].value,
                                                   self.__listOfFilesIButton)
                self.__createFile_iButton(sKeyData,
                                          sIButtonKeyType,
                                          sFileName+".ibtn")
            nStartRow += 1


if __name__ == '__main__':
    converter = Keys_Xlsx2FlipperFiles(FILE_NAME)
    print("Convetring xlsx to Flipper Files...")
    nErrs = converter.convert()
    if nErrs == 0:
        print("Succesfully converted!")
    else:
        print(f"Converted with {nErrs} errors!")

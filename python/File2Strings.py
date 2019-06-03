# -*- coding:utf-8 -*-

from optparse import OptionParser
from XlsFileUtil import XlsFileUtil
from XmlFileUtil import XmlFileUtil
from StringsFileUtil import StringsFileUtil
from Log import Log
import os
import time


def addParser():
    parser = OptionParser()

    parser.add_option("-f", "--fileDir",
                      help="Xls files directory.",
                      metavar="fileDir")

    parser.add_option("-t", "--targetDir",
                      help="The directory where the strings files will be saved.",
                      metavar="targetDir")

    parser.add_option("-e", "--excelStorageForm",
                      type="string",
                      default="multiple",
                      help="The excel(.xls) file storage forms including single(single file), multiple(multiple files), default is multiple.",
                      metavar="excelStorageForm")

    parser.add_option("-a", "--additional",
                      help="additional info.",
                      metavar="additional")

    (options, args) = parser.parse_args()
    Log.info("options: %s, args: %s" % (options, args))

    return options


def convertFromSingleForm(options, fileDir, targetDir):
    for _, _, filenames in os.walk(fileDir):
        xlsFilenames = [fi for fi in filenames if fi.endswith(".xls")]
        for file in xlsFilenames:
            xlsFileUtil = XlsFileUtil(fileDir+"/"+file)
            table = xlsFileUtil.getTableByIndex(0)
            firstRow = table.row_values(0)
            keys = table.col_values(0)
            del keys[0]

            for index in range(len(firstRow)):
                if index <= 0:
                    continue
                languageName = firstRow[index]
                values = table.col_values(index)
                del values[0]
                StringsFileUtil.writeToFile(
                    keys, values, targetDir + "/"+languageName+".lproj/", file.replace(".xls", "")+".strings", options.additional)
    print "Convert %s successfully! you can see strings file in %s" % (
        fileDir, targetDir)

# 获取lingoes里支持的所有单词
def convertFromFileFrom(options, fileDir, targetDir):
    iosDestFilePath = targetDir + "/" + "Lingoes"
    iosFileManager = open(iosDestFilePath, "wb")
    iosFileManager.write("[\n")
    for _, dirs, _ in os.walk(fileDir):
        for dirName in dirs:
            fileNameDir = fileDir + dirName
            print "====== %s" % (fileNameDir)
            for _, _, filenames in os.walk(fileNameDir):
                for file in filenames:
                    content = "\"%s\",\n" % (file.replace(".mp3", ""))
                    iosFileManager.write(content)
    iosFileManager.write("\n]")
    if options.additional is not None:
        iosFileManager.write(options.additional)
    iosFileManager.close()
    print "Convert %s successfully! you can see strings file in %s" % (fileDir, targetDir)


def convertFromMultipleForm(options, fileDir, targetDir):
    for _, _, filenames in os.walk(fileDir):
        xlsFilenames = [fi for fi in filenames if fi.endswith(".xls")]
        for file in xlsFilenames:
            xlsFileUtil = XlsFileUtil(fileDir+"/"+file)
            langFolderPath = targetDir + "/" + file.replace(".xls", "")
            if not os.path.exists(langFolderPath):
                os.makedirs(langFolderPath)

            for sheet in xlsFileUtil.getAllTables():
                iosDestFilePath = langFolderPath + "/" + file.replace(".xls", "")
                iosFileManager = open(iosDestFilePath, "wb")
                iosFileManager.write("[\n")
                for row in sheet.get_rows():
                  # 换行符替换为空格
                    content = "{\"" + row[0].value + "\" " + \
                      ": " + "\"" + row[1].value.replace("\n", " ") + "\"},\n"
#                    content = row[0].value + "@" + row[1].value.replace("\n", "; ") + "\n"
                    iosFileManager.write(content)
                iosFileManager.write("\n]")
                if options.additional is not None:
                    iosFileManager.write(options.additional)
                iosFileManager.close()
    print "Convert %s successfully! you can see strings file in %s" % (
        fileDir, targetDir)


def startConvert(options):
    fileDir = options.fileDir
    targetDir = options.targetDir

    print "Start converting"

    targetDir = targetDir + "/xls-files-to-strings_" + \
        time.strftime("%Y%m%d_%H%M%S")
    if not os.path.exists(targetDir):
        os.makedirs(targetDir)

    convertFromFileFrom(options, fileDir, targetDir)


def main():
    options = addParser()
    startConvert(options)


main()

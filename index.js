$(function() {
  $('#excelFile').change(function(parentEvent) {
    let files = parentEvent.target.files;
    console.log("===files::")
    console.log(files)
    let keyI = '编号';
    let keyJ = '名字';
    let keyO = '时间';

    let fileReader = new FileReader();

    let getExcelList = [];
    fileReader.onload = function(childEvent) {
      let excelBinaryData;
      // 读取上传的excel文件
      try {
        let excelData = childEvent.target.result;
        excelBinaryData = XLSX.read(excelData, {
          type: 'binary'
        });
      } catch (parentEvent) {
        console.log('该文件类型不能识别');
        return;
      }

      // 获取excell所有元素
      let sheetIndex = 0;
      let headNameList = [];
      let allHeadName = new Set();
      for (let sheet in excelBinaryData.Sheets) {
        if (excelBinaryData.Sheets.hasOwnProperty(sheet)) {
          let excelSheet = XLSX.utils.sheet_to_json(excelBinaryData.Sheets[sheet]);
          getExcelList[sheetIndex] = excelSheet;

          // 获取excel表名
          let elementIndex = 0;
          let elementList = [];
          for (let headIndex in excelBinaryData.Sheets[sheet]) {
            if (headIndex.indexOf(1) != -1) {
              elementList[elementIndex] = excelBinaryData.Sheets[sheet][headIndex].w;
              allHeadName.add(excelBinaryData.Sheets[sheet][headIndex].w);
              elementIndex++;
            }
          }
          headNameList[sheetIndex] = elementList;
          sheetIndex++;
        }
      }
      headNameList[headNameList.length] = Array.from(allHeadName);

      let elementNameList = [];

      let getCompareSameOne = [];
      let getCompareSameTwo = [];

      // 找出可以合并的数据
      for (let sheetI = 0; sheetI < getExcelList[0].length; sheetI++) {
        for (let sheetJ = 0; sheetJ < getExcelList[1].length; sheetJ++) {
          let excelOne = getExcelList[0][sheetI];
          let excelTwo = getExcelList[1][sheetJ];

          // 如果三个key都相同或者有两个key相同
          if (compareCondition(excelOne, excelTwo)) {
            getCompareSameOne[getCompareSameOne.length] = excelOne;
            getCompareSameTwo[getCompareSameTwo.length] = excelTwo;

            let headElement = headNameList[headNameList.length - 1];
            let newLength = headElement.length;

            let elementNameMap = {};
            for (let newIndex = 0; newIndex < newLength; newIndex++) {
              let elementName = headElement[newIndex];

              // 如果第二个sheet里面的值不为空，则值为第二各sheet里的值
              // 否则为第一个sheet里面的值
              if (excelTwo[elementName]) {
                elementNameMap[elementName] = excelTwo[elementName];
              } else {
                elementNameMap[elementName] = excelOne[elementName];
              }
            }
            elementNameList[elementNameList.length] = elementNameMap;
          }
        }
      }

      elementNameList = findCompareNotSame(getExcelList[0], getCompareSameOne, elementNameList);
      elementNameList = findCompareNotSame(getExcelList[1], getCompareSameTwo, elementNameList);
      getExcelList[getExcelList.length] = elementNameList;

      download(getExcelList);
    };

    function download(getExcelList) {
      console.log("====getExcelList:")
      console.log(getExcelList);
      const newSheet = {
        SheetNames: ['Sheet1', 'Sheet2', 'Sheet3'],
        Sheets: {},
        Props: {}
      };
      const sheetDownloadType = { bookType: 'xlsx', bookSST: false, type: 'binary' };

      newSheet.Sheets['Sheet1'] = XLSX.utils.json_to_sheet(getExcelList[0]);
      newSheet.Sheets['Sheet2'] = XLSX.utils.json_to_sheet(getExcelList[1]);
      newSheet.Sheets['Sheet3'] = XLSX.utils.json_to_sheet(getExcelList[2]);
      saveAs(
        new Blob(
          [
            stringToArrayBuffer(XLSX.write(newSheet, sheetDownloadType))
          ], {
            type: "application/octet-stream"
          }
        ),
        files[0].name
      );
    }

    function stringToArrayBuffer(data) {
      if (typeof ArrayBuffer !== 'undefined') {
        let arrayBuffer = new ArrayBuffer(data.length);
        let unitArray = new Uint8Array(arrayBuffer);
        for (let unitI = 0; unitI != data.length; unitI++) {
          unitArray[unitI] = data.charCodeAt(unitI) & 0xFF;
        }
        return arrayBuffer;
      } else {
        let arrayBuffer = new Array(data.length);
        for (let bufferI = 0; bufferI != data.length; bufferI++) {
          arrayBuffer[bufferI] = data.charCodeAt(bufferI) & 0xFF;
        }
        return arrayBuffer;
      }
    }

    // 找出sheet表不能合并的数据
    function findCompareNotSame(getExcelListNumber, getCompareSame, elementNameList) {
      for (let sheetI = 0; sheetI < getExcelListNumber.length; sheetI++) {
        let notSameCount = 0;
        let excelOne = getExcelListNumber[sheetI];
        for (let sheetSame = 0; sheetSame < getCompareSame.length; sheetSame++) {
          let excelSame = getCompareSame[sheetSame];
          if (compareCondition(excelOne, excelSame)) {
            notSameCount++;
          }
        }

        if (!notSameCount) {
          elementNameList[elementNameList.length] = excelOne;
        }
      }
      return elementNameList;
    }

    function compareCondition(excelOne, excelTwo) {
      return excelOne[keyI] == excelTwo[keyI] && excelOne[keyJ] == excelTwo[keyJ] && excelOne[keyO] == excelTwo[keyO]
        || excelOne[keyI] == excelTwo[keyI] && excelOne[keyJ] == excelTwo[keyJ] && (excelOne[keyO] == '' || excelTwo[keyO] == '')
        || excelOne[keyI] == excelTwo[keyI] && excelOne[keyO] == excelTwo[keyO] && (excelOne[keyJ] == '' || excelTwo[keyJ] == '')
        || excelOne[keyJ] == excelTwo[keyJ] && excelOne[keyO] == excelTwo[keyO] && (excelOne[keyI] == '' || excelTwo[keyI] == '');
    }

    function saveAs(content, fileName) {
      let clickDiv = document.createElement("a");
      clickDiv.download = fileName || "下载";
      clickDiv.href = URL.createObjectURL(content);
      clickDiv.click();
      setTimeout(function () {
        URL.revokeObjectURL(content);
      }, 100);
    }

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });
})

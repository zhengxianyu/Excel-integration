$(function() {
  $('#excelFile').change(function(parentEvent) {
    let files = parentEvent.target.files;
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
      saveAs(new Blob([s2ab(XLSX.write(newSheet, sheetDownloadType))], { type: "application/octet-stream" }), "合成数据表格" + '.' + (sheetDownloadType.bookType=="biff2"?"xls":sheetDownloadType.bookType));
    }

    function s2ab(s) {
      if (typeof ArrayBuffer !== 'undefined') {
        let buf = new ArrayBuffer(s.length);
        let view = new Uint8Array(buf);
        for (let i = 0; i != s.length; ++i) {
          view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
      } else {
        let buf = new Array(s.length);
        for (let i = 0; i != s.length; ++i) {
          buf[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
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
      return excelOne['编号'] == excelTwo['编号'] && excelOne['名字'] == excelTwo['名字'] && excelOne['时间'] == excelTwo['时间']
        || excelOne['编号'] == excelTwo['编号'] && excelOne['名字'] == excelTwo['名字'] && (excelOne['时间'] == '' || excelTwo['时间'] == '')
        || excelOne['编号'] == excelTwo['编号'] && excelOne['时间'] == excelTwo['时间'] && (excelOne['名字'] == '' || excelTwo['名字'] == '')
        || excelOne['名字'] == excelTwo['名字'] && excelOne['时间'] == excelTwo['时间'] && (excelOne['编号'] == '' || excelTwo['编号'] == '');
    }

    function saveAs(obj, fileName) {//当然可以自定义简单的下载文件实现方式 
      let tmpa = document.createElement("a");
      tmpa.download = fileName || "下载";
      tmpa.href = URL.createObjectURL(obj); //绑定a标签
      tmpa.click(); //模拟点击实现下载
      setTimeout(function () { //延时释放
        URL.revokeObjectURL(obj); //用URL.revokeObjectURL()来释放这个object URL
      }, 100);
    }

    // 以二进制方式打开文件
    fileReader.readAsBinaryString(files[0]);
  });
})

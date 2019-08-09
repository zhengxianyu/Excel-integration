$(function() {
  $('#excelFile').change(function(parentEvent) {
    let files = parentEvent.target.files;
    let fileReader = new FileReader();

    fileReader.onload = function(childEvent) {
      let excelBinaryData, getExcelList = [];
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
          console.log("===excelBinaryData.Sheets[sheet]::")
          console.log(excelBinaryData.Sheets[sheet])
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
      console.log("==headNameList:::")
      console.log(headNameList)

      let newAllSameIndex = 0;
      let elementNameList = [];

      for (let sheetI = 0; sheetI < getExcelList[0].length; sheetI++) {
        for (let sheetJ = 0; sheetJ < getExcelList[1].length; sheetJ++) {
          let excelOne = getExcelList[0][sheetI];
          let excelTwo = getExcelList[1][sheetJ];

          // 如果三个key都相同或者有两个key相同
          if (compareCondition(excelOne, excelTwo)) {
            let headElement = headNameList[headNameList.length - 1];
            let newLength = headElement.length;
            
            let elementNameMap = [];
            for (let newIndex = 0; newIndex < newLength; newIndex++) {
              let elementName = headElement[newIndex];
              if (excelTwo[elementName]) {
                elementNameMap[elementName] = excelTwo[elementName];
              } else {
                elementNameMap[elementName] = excelOne[elementName];
              }
            }
            elementNameList[newAllSameIndex] = elementNameMap;
            newAllSameIndex++;
          }
        }
      }
      getExcelList[getExcelList.length] = elementNameList;
      console.log("====getExcelList:")
      console.log(getExcelList);
    };

    function compareCondition(excelOne, excelTwo) {
      return excelOne['编号'] == excelTwo['编号'] && excelOne['名字'] == excelTwo['名字'] && excelOne['时间'] == excelTwo['时间']
        || excelOne['编号'] == excelTwo['编号'] && excelOne['名字'] == excelTwo['名字'] && (excelOne['时间'] == '' || excelTwo['时间'] == '')
        || excelOne['编号'] == excelTwo['编号'] && excelOne['时间'] == excelTwo['时间'] && (excelOne['名字'] == '' || excelTwo['名字'] == '')
        || excelOne['名字'] == excelTwo['名字'] && excelOne['时间'] == excelTwo['时间'] && (excelOne['编号'] == '' || excelTwo['编号'] == '');
    }
    // 以二进制方式打开文件
    // fileReader.readAsBinaryString(files[0]);
  });
})

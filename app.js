const xlsx = require('xlsx'),
      fs   = require('fs'),
      path = require('path');


//takes an array of objects and returns another array of objects with a certain property removed
function property_remover(data, property){
  let new_data = data.map(record => {
    if(record[property]){
      delete(record[property]);
    }
    return record;
  })
  return new_data;
}



//takes the path of a certain excel workbook path and returns the data in its first worksheet
function readFileToJson(workbook_path){
  let wb = xlsx.readFile(workbook_path),
      first_tab = wb.SheetNames[0],
      ws = wb.Sheets[first_tab],
      data = xlsx.utils.sheet_to_json(ws);
      data = property_remover(data, 'مسلسل');
  return data;
}

function extension_validator(file){
  let file_ext = path.parse(file).ext;
  if( ((file_ext === '.xls')||(file_ext === '.xlsx')) && file[0] !== "~"){
    return true;
  }
  return false;
}

function arr_comparator(agg_arr, obj, code) {
  for (i = 0; i < agg_arr.length; i++){
    agg_obj = agg_arr[i];
    if(obj[code] === agg_obj[code]){
      return i;
    }
  }
  return false;
}

//takes an array of duplicate objects and returns an array with its duplicate objects reduced to one with a counter holding the number of duplications
function data_consolidator(attendance_data){
    let aggregate_att_arr = [],
        mod_att_data = attendance_data.map( obj => {
          return {...obj, counter:1};
        }),
        code = ['الكود'];
        mod_att_data.forEach(obj => {
          let object_exists = arr_comparator(aggregate_att_arr, obj, code);
          if(object_exists){
            agg_obj = aggregate_att_arr[object_exists];
            agg_obj.counter++;
          }
          else{
            aggregate_att_arr.push(obj);
          }
        })
    return aggregate_att_arr;
}



//takes the path of the folder containing our excel files that we want to merge, and returns an array containing their data combined and reduced (using data_consolidator) in json format
function files_merger_generic(sourceDir){
  let dir_path = path.join(__dirname, sourceDir),
      files = fs.readdirSync(dir_path).filter(extension_validator),
      combinedFilesArray = [];
  files.forEach(file => {
      fileData = readFileToJson(path.join(__dirname, sourceDir, file));
      combinedFilesArray = combinedFilesArray.concat(fileData);
  })
  let reducedFilesArray = data_consolidator(combinedFilesArray);
  return reducedFilesArray;
}

//takes the path of the folder containing our excel files that we want to merge, and returns an array containing their data combined in json format(1 array of monthly payments and 1 array of other payments)
function files_merger_cash(sourceDir){
  let dir_path = path.join(__dirname, sourceDir),
      files = fs.readdirSync(dir_path).filter(extension_validator),
      val = 'اشتراك',
      combined_monthly = [],
      combined_other = [];
  files.forEach(file => {
      fileData = readFileToJson(path.join(__dirname, sourceDir, file));
      monthly_fees = fileData.filter(obj => obj['بيان'] === val);
      other_fees = fileData.filter(obj => obj['بيان'] !== val);
      combined_monthly = combined_monthly.concat(monthly_fees);
      combined_other = combined_other.concat(other_fees);
  })
  return [combined_monthly, combined_other];
}

//takes the path of the payments folder containing the total and cash excel files each in its own folder, and returns an array of objects containing the students that haven't paid and are not terminated
function net_cash_calculator(source_dir){
  let dir_path_cash = path.join(__dirname, source_dir, 'cash'),
      dir_path_total = path.join(__dirname, source_dir, 'total'),
      cash_file = fs.readdirSync(dir_path_cash).filter(extension_validator)[0],
      total_file = fs.readdirSync(dir_path_total).filter(extension_validator)[0],
      file_data_1 = readFileToJson(path.join(dir_path_cash, cash_file)),
      file_data_2 = readFileToJson(path.join(dir_path_total, total_file)),
      unique_data = [...file_data_2],
      std_code = 'الكود';
      for ( i = 0; i < file_data_2.length; i++){
        for (j = 0; j < file_data_1.length; j++){
          code_cash = file_data_1[j][std_code];
          code_total = file_data_2[i][std_code];
          if (code_cash === code_total){
            unique_data.splice(unique_data.findIndex( obj => obj[std_code] === code_total), 1);
            break;
          }
        }
      }
      let active_unique_data = unique_data.filter(std => std['الموقف'] !== 'Termination');
    return active_unique_data;
}


//takes the path of the folder containing our excel files that we want to merge, and returns a new file containing all of their data combined
function merged_file_creator(sourceDir, data_array, file_name){
  let newWB = xlsx.utils.book_new(),
      newWS = xlsx.utils.json_to_sheet(data_array);
      xlsx.utils.book_append_sheet(newWB, newWS, "Merged Data");
      xlsx.writeFile(newWB, file_name);
}

function directory_creator(folder_name){
  let sourceDir = 'excel_files',
      dir = path.join(sourceDir, folder_name);
  return dir;
}

let att_src_dir = directory_creator('attendance');
let trans_src_dir = directory_creator('transforms');
let cash_src_dir = directory_creator('cash');
let payments_src_dir = directory_creator('payments');


merged_file_creator(att_src_dir, files_merger_generic(att_src_dir), "اجمالي حضور.xlsx");
merged_file_creator(trans_src_dir, files_merger_generic(trans_src_dir), "اجمالي تحويلات.xlsx");
merged_file_creator(cash_src_dir, files_merger_cash(cash_src_dir)[0], "اجمالي يومية اشتراكات.xlsx");
merged_file_creator(cash_src_dir, files_merger_cash(cash_src_dir)[1], "اجمالي يومية اخري.xlsx");
merged_file_creator(payments_src_dir, net_cash_calculator(payments_src_dir), 'غير مدفوع.xlsx');

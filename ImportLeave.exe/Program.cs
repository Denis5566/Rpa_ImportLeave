using NLog;
using NLog.Common;
using OfficeOpenXml;
using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;

namespace LeaveImport
{
    class Program
    {
        private static Logger logger;
        private static readonly string RunId = Guid.NewGuid().ToString("N");
        private static readonly string logFolder =
      Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LOG");
        // A：請假資料、B：公出/出差資料
        private static string filetype = "A";

        #region 寫LOGFunction
        private static void WriteLog(string message, bool isError = false)
        {
            // 如果 LOG 資料夾不存在就建立
            if (!Directory.Exists(logFolder))
            {
                try
                {
                    Directory.CreateDirectory(logFolder);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }

            string logMessage =
                DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
                + " "
                + message;

            Console.WriteLine(logMessage);

            // 一般 Log 永遠都寫
            string normalLog = Path.Combine(logFolder, "LeaveImport.log");
            File.AppendAllText(normalLog, logMessage + Environment.NewLine);

            // 如果是錯誤，再另外寫一份 Error Log
            if (isError)
            {

                string errorLog = Path.Combine(logFolder, "ErrMessage.log");
                File.AppendAllText(
                    errorLog,
                    logMessage + Environment.NewLine);
            }
        }
        #endregion

        #region 主程式
        static void Main(string[] args)
        {
            WriteLog("程式開始" + RunId);

           

            try
            {
                ExcelPackage.License.SetNonCommercialPersonal("Denis");

                // 先找請假
                string leavePath =ConfigurationManager.AppSettings["LeaveExcelPath"];           
                leavePath = leavePath.Replace("{Date}", DateTime.Today.ToString("yyyy_MM_dd"));

                // 再找公出/出差
                string tripPath = ConfigurationManager.AppSettings["TripExcelPath"];
                tripPath = tripPath.Replace("{Date}", DateTime.Today.ToString("yyyy-MM-dd"));

                if (string.IsNullOrEmpty(leavePath) || string.IsNullOrEmpty(tripPath))
                {
                    WriteLog("ERROR：未設定 ExcelPath 或 TripExcelPath");
                    return;
                }
                string filePath = "";
                if (File.Exists(leavePath))
                {
                    filePath = leavePath;
                    filetype = "A";
                    WriteLog("使用請假資料：" + Path.GetFileName(filePath));
                }
                else if (File.Exists(tripPath))
                {
                    filePath = tripPath;
                    filetype = "B";
                    WriteLog("使用公出/出差資料：" + Path.GetFileName(filePath));
                }
                else
                {
                    WriteLog("ERROR：找不到 Excel 檔");
                    WriteLog("請假：" + leavePath);
                    WriteLog("公出：" + tripPath);
                    return;
                }

                int successCount = 0;
                int failCount = 0;

                using (ExcelPackage package =
                    new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet ws =
                        package.Workbook.Worksheets[0];

                    int rowCount = ws.Dimension.End.Row;
                    int skipCount = 0;

                    WriteLog("共讀取 " + (rowCount - 1) + " 筆資料");
                    WriteLog("開始刪除 D_SOURCE='" + filetype + "' 舊資料");
                    DeleteOldData();
                    WriteLog("刪除完成");

                    //狀態的位置
                    int COL_STATUS = filetype == "B" ? 21 : 16;
                    string formStatus = "";
                    for (int row = 2; row <= rowCount; row++)
                    {
                        try
                        {
                            formStatus =  ws.Cells[row, COL_STATUS].Text.Trim();
                            if (formStatus != "已核准")
                            {
                                skipCount++;

                                string rowData = "";

                                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                                {
                                    rowData += $"[{ws.Cells[1, col].Text}]={ws.Cells[row, col].Text} ";
                                }

                                WriteLog(
                                    $"略過原因：表單狀態不是「已核准」(目前狀態：{formStatus})，Row={row}，{rowData}");

                                continue;
                            }

                            if (filetype == "B")
                            {
                                InsertData(
                                    ws.Cells[row, 4].Text,   // 工號
                                    ws.Cells[row, 8].Text,   // 出差起
                                    ws.Cells[row, 9].Text,   // 出差迄
                                    ws.Cells[row, 2].Text,   // 部門
                                    ws.Cells[row, 5].Text,   // 中文名
                                    ws.Cells[row, 6].Text,   // 英文名
                                    ws.Cells[row, 12].Text,  // 類型(出差/公出)
                                    ws.Cells[row, 11].Text,  // 時數
                                    ws.Cells[row, 10].Text,  // 日數
                                    ws.Cells[row, 16].Text,  // 代理人
                                    ws.Cells[row, 22].Text   // 申請日期
                                );
                            }
                            else
                            {
                                InsertData(
                                    ws.Cells[row, 2].Text,   //工號
                                    ws.Cells[row, 5].Text,   //請假起
                                    ws.Cells[row, 6].Text,   //請假迄
                                    ws.Cells[row, 1].Text,   //部門
                                    ws.Cells[row, 3].Text,   //中文名
                                    ws.Cells[row, 4].Text,   //英文名
                                    ws.Cells[row, 7].Text,   //假別
                                    ws.Cells[row, 8].Text,   //時數
                                    ws.Cells[row, 10].Text,  //日數
                                    ws.Cells[row, 11].Text,  //代理人
                                    ws.Cells[row, 17].Text   //申請日期
                                );
                            }
                            successCount++;
                            
                            string empNo = filetype == "B"
                            ? ws.Cells[row, 4].Text
                            : ws.Cells[row, 2].Text;
                            WriteLog($"成功 工號:{empNo}");
                   
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            string empNo = filetype == "B"
                            ? ws.Cells[row, 4].Text
                            : ws.Cells[row, 2].Text;

                            WriteLog(
                                $"失敗 Row={row} 工號:{empNo}", true);
                         
                            WriteLog(ex.ToString(),true);
                        }
                    }
                }
             
                WriteLog(
                    $"匯入完成 成功:{successCount} 失敗:{failCount}");

                

            }
            catch (Exception ex)
            {
                WriteLog("程式異常終止",true);

                WriteLog(ex.ToString(),true);
            }
            finally
            {
                //if (File.Exists(filePath))
                //{
                //    File.Delete(filePath);
                //    WriteLog("已刪除 Excel：" + filePath);
                //}

                WriteLog("程式結束");
                Process.Start("explorer.exe", logFolder);           
              //  File.Create(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ImportLeave.end")).Dispose();
            }
        }
        #endregion

        #region 插入請假資料列
        static void InsertData(
            string empId,
            string startTime,
            string endTime,
            string deptName,
            string chName,
            string enName,
            string absType,
            string absHour,
            string absDay,
            string proxy,
            string logTime)
        {
            string connStr =
                ConfigurationManager.ConnectionStrings["DB"].ConnectionString;

            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();

                string sql = @"
INSERT INTO ZHR_ABSENCE
(
 EMP_ID,
 START_TIME,
 END_TIME,
 DEPT_NAME,
 CH_NAME,
 EN_NAME,
 ABS_TYPE,
 ABS_HOUR,
 ABS_DAY,
 JOB_PROXY,

 D_SOURCE
)
VALUES
(
 @EMP_ID,
 @START_TIME,
 @END_TIME,
 @DEPT_NAME,
 @CH_NAME,
 @EN_NAME,
 @ABS_TYPE,
 @ABS_HOUR,
 @ABS_DAY,
 @JOB_PROXY,
 @filetype
)";

                using (SqlCommand cmd =
                    new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@EMP_ID", empId);

                    cmd.Parameters.AddWithValue(
                        "@START_TIME",
                        DateTime.Parse(startTime)
                            .ToString("yyyy/MM/dd HH:mm"));

                    cmd.Parameters.AddWithValue(
                        "@END_TIME",
                        DateTime.Parse(endTime)
                            .ToString("yyyy/MM/dd HH:mm"));

                    cmd.Parameters.AddWithValue("@DEPT_NAME", deptName);
                    cmd.Parameters.AddWithValue("@CH_NAME", chName);
                    cmd.Parameters.AddWithValue("@EN_NAME", enName);
                    cmd.Parameters.AddWithValue("@ABS_TYPE", absType);
                    cmd.Parameters.AddWithValue("@filetype", filetype);
                    cmd.Parameters.AddWithValue(
                        "@ABS_HOUR",
                        Convert.ToDecimal(absHour));

                    cmd.Parameters.AddWithValue(
                        "@ABS_DAY",
                        Convert.ToDecimal(absDay));

                    cmd.Parameters.AddWithValue("@JOB_PROXY", proxy);

                    //cmd.Parameters.AddWithValue(
                    //    "@LOG_TIME",
                    //    DateTime.Parse(logTime)
                    //        .ToString("yyyy/MM/dd HH:mm"));

                    cmd.ExecuteNonQuery();
                }
            }
        }
        #endregion

        #region 刪除舊的請假/公出資料列
        static void DeleteOldData()
        {
            string connStr =
                ConfigurationManager.ConnectionStrings["DB"].ConnectionString;
         
            using (SqlConnection conn = new SqlConnection(connStr))
            {
                conn.Open();
               // WriteLog("ConnectionString：" + conn.ConnectionString);
                string sql = @"
DELETE FROM ZHR_ABSENCE
WHERE D_SOURCE = @D_SOURCE";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.AddWithValue("@D_SOURCE", filetype);

                    int count = cmd.ExecuteNonQuery();

                    WriteLog($"已刪除 {count} 筆 D_SOURCE='{filetype}' 資料");
                }
            }
        }
        #endregion

    }
}
       
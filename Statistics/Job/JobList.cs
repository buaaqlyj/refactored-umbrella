using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using Statistics.Log;
using Statistics.SuperDog;

namespace Statistics.Job
{
    public class JobList
    {
        private List<Job> jobList;
        private JobParameterStruct jobParam;
        private bool isStopping;
        private int jobCount;

        public JobList(FileInfo[] files, JobParameterStruct jobParam, Person person)
        {
            this.jobList = new List<Job>();
            this.jobParam = jobParam;
            this.isStopping = false;
            switch (jobParam.ActionType)
            {
                case 0:
                    foreach (FileInfo item in files)
                    {
                        AddJob(new GeneratingFormJob(item.FullName, jobParam, person));
                    }
                    break;
                case 1:
                    foreach (FileInfo item in files)
                    {
                        AddJob(new GeneratingCertificateJob(item.FullName, jobParam, person));
                    }
                    break;
                default:
                    throw new Exception("生成任务时无法识别选择的任务类型！");
            }
            this.jobCount = jobList.Count;
        }

        #region Public Method
        public void AddJob(Job job)
        {
            if (job.Equals(null))
            {
                throw new Exception("要加入的任务为空！");
            }
            jobList.Add(job);
        }

        public void ClearJob()
        {
            this.jobList.Clear();
            this.jobParam = null;
            this.isStopping = false;
        }

        public void DoWork()
        {
            LogHelper.StartNewTask(jobCount);

            if (jobCount > 0)
            {
                foreach (Job item in jobList)
                {
                    if (!item.FileName.StartsWith(@"~$"))
                    {
                        LogHelper.StartNewJob(item.FileName);

                        try
                        {
                            item.DoTheJob();
                        }
                        catch (Exception ex)
                        {
                            LogHelper.AddLog(@"* 文件名称：" + item.FullName, true);
                            LogHelper.AddLog(@"* 异常消息：" + ex.Message, true);
                            LogHelper.AddLog(@"* 异常方法：" + ex.TargetSite, true);

                            LogHelper.AddProblemFilesAndReset(item.FullName);
                        }
                        finally
                        {
                            LogHelper.FinishOneJob();
                        }
                    }
                    if (IsStopping)
                    {
                        break;
                    }
                }
            }
            else
            {
                throw new Exception(@"输入文件夹没有找到待处理的文件");
            }
        }
        #endregion

        #region Property
        public bool IsStopping
        {
            get
            {
                return isStopping;
            }
            set
            {
                isStopping = value;
            }
        }
        #endregion
    }
}

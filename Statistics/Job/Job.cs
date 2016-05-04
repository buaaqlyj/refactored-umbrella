using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using Statistics.SuperDog;

namespace Statistics.Job
{
    public abstract class Job
    {
        protected JobParameterStruct jobParam = null;
        protected string filePath;
        protected string fileName;
        protected Person person = null;

        public Job(string filePath, JobParameterStruct jobParam, Person person)
        {
            this.jobParam = jobParam;
            this.filePath = filePath;
            this.fileName = Path.GetFileName(filePath);
            this.person = person;
        }

        public abstract void DoTheJob();

        public abstract JobType JobType { get; }

        public string FileName
        {
            get
            {
                return fileName;
            }
        }

        public string FullName
        {
            get
            {
                return filePath;
            }
        }
    }
}

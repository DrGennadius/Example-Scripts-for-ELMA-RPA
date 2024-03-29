using System;
using System.IO;
using System.Collections.Generic;
using ELMA.RPA.SDK.Commons.Params;

namespace ELMA.RPA.Scripts
{
    /// <summary>
    /// Контекст процесса
    /// <summary>
    public class Context
    {
        // <value>StringParam</value>
        public String stringparam
        {
            get;
            set;
        }

        // <value>FileParam</value>
        public FileInfo fileparam
        {
            get;
            set;
        }

        // <value>DateTimeParam</value>
        public Nullable<DateTime> datetimeparam
        {
            get;
            set;
        }

        // <value>BooleanParam</value>
        public Nullable<Boolean> booleanparam
        {
            get;
            set;
        }

        // <value>DoubleParam</value>
        public Nullable<Double> doubleparam
        {
            get;
            set;
        }

        // <value>integerparam</value>
        public Nullable<Int64> integerparam
        {
            get;
            set;
        }

        // <value>TableParam</value>
        public RpaTable tableparam
        {
            get;
            set;
        }

        // <value>StructParam</value>
        public RpaStructure structparam
        {
            get;
            set;
        }

        // <value>StringList</value>
        public List<String> stringlist
        {
            get;
            set;
        }

        // <value>FileList</value>
        public List<FileInfo> filelist
        {
            get;
            set;
        }

        // <value>DateTimeList</value>
        public List<Nullable<DateTime>> datetimelist
        {
            get;
            set;
        }

        // <value>BooleanList</value>
        public List<Nullable<Boolean>> booleanlist
        {
            get;
            set;
        }

        // <value>DoubleList</value>
        public List<Nullable<Double>> doublelist
        {
            get;
            set;
        }

        // <value>IntegerList</value>
        public List<Nullable<Int64>> integerlist
        {
            get;
            set;
        }
    }
}
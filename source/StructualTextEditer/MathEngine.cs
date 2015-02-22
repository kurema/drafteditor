using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StructualTextEditer
{
    class MathEngine
    {
        public List<OperationBasic> Operations = new List<OperationBasic>();

        public MathEngine()
        {
            AddOperation(new OperationBasic(new OperationBasic.DoubleArgOperation(BasicOperations.Add),"+"));
            AddOperation(new OperationBasic(new OperationBasic.DoubleArgOperation(BasicOperations.Sub), "-"));
            AddOperation(new OperationBasic(new OperationBasic.DoubleArgOperation(BasicOperations.Times), "*"));
            AddOperation(new OperationBasic(new OperationBasic.DoubleArgOperation(BasicOperations.Divide), "/"));
        }
        
        public void AddOperation(OperationBasic op)
        {
            Operations.Add(op);
        }


        public static class BasicOperations
        {
            public static Int64 Add(Int64 a,Int64 b){return a+b;}
            public static Int64 Sub(Int64 a, Int64 b) { return a - b; }
            public static Int64 Times(Int64 a, Int64 b) { return a * b; }
            public static Int64 Divide(Int64 a, Int64 b) { return a / b; }
        }

        public class Formula
        {
            public OperationBasic Operation;
            public List<Formula> Args;

            public static implicit operator long(Formula v) {
                if (v is Value) { return v; }
                else
                {
                    long[] arg = new long[v.Args.Count];
                    for (int i = 0; i < v.Args.Count; i++)
                    {
                        arg[i] = v.Args[i];
                    }
                    return v.Operation.Operate(arg);
                }
            }
            
            public Formula Calculate()
            {//細かいところチェック
                try
                {
                    return new Value(this);
                }
                catch//型変換失敗をキャッチ
                {
                    return this;//ToDo:式をまとめる。
                }
            }
        }

        public class Value:Formula
        {
            public long Number;

            public static implicit operator long(Value v) { return v.Number; }
            public static implicit operator Value(long v) { return new Value(v); }

            public long Calculate()
            {
                return Number;
            }

            public Value()
            {
                Number = 1;
            }

            public Value(long v) { Number = v; }
        }

        public class OperationBasic
        {
            public delegate Int64 NoArgOperation();
            public delegate Int64 SingleArgOperation(Int64 arg);
            public delegate Int64 DoubleArgOperation(Int64 arg1,Int64 arg2);
            public delegate Int64 MultipleArgOperation(params Int64[] args);
            private Dictionary<DisplayType,string> display;
            private Dictionary<DisplayType, string> displayOnParse;
            private Object Operation;//上に定義されているdelegateのいずれか(*Operation)。

            public enum DisplayType
            {
                basic,tex,MSOffice,MathML
            }

            public OperationBasic()
            {
                Operation = null;
            }

            public OperationBasic(NoArgOperation op) { SetDelagate(op); }
            public OperationBasic(SingleArgOperation op) { SetDelagate(op); }
            public OperationBasic(DoubleArgOperation op) { SetDelagate(op); }
            public OperationBasic(MultipleArgOperation op) { SetDelagate(op); }

            public OperationBasic(NoArgOperation op, string text,DisplayType dt=DisplayType.basic) : this(op) { SetDisplay(text,dt); }
            public OperationBasic(SingleArgOperation op, string text, DisplayType dt = DisplayType.basic) : this(op) { SetDisplay(text, dt); }
            public OperationBasic(DoubleArgOperation op, string text, DisplayType dt = DisplayType.basic) : this(op) { SetDisplay(text, dt); }
            public OperationBasic(MultipleArgOperation op, string text, DisplayType dt = DisplayType.basic) : this(op) { SetDisplay(text, dt); }

            public Int64 Operate(params Int64[] args)
            {
                if (Operation == null)
                {
                    throw new NotImplementedException();
                }
                else if (Operation is NoArgOperation)
                {
                    return ((NoArgOperation)Operation)();
                }
                else if (Operation is SingleArgOperation)
                {
                    return ((SingleArgOperation)Operation)(args[0]);
                }
                else if (Operation is DoubleArgOperation)
                {
                    return ((DoubleArgOperation)Operation)(args[0],args[1]);
                }
                else if (Operation is MultipleArgOperation)
                {
                    return ((MultipleArgOperation)Operation)(args);
                }
                else
                {
                    return OriginalOperation(args);
                }
            }

            protected Int64 OriginalOperation(params Int64[] args){
                throw new NotImplementedException();
            }

            public OperationBasic SetDelagate(object op)
            {
                if (op is SingleArgOperation || op is DoubleArgOperation || op is MultipleArgOperation || op is NoArgOperation)
                {
                    Operation = op;
                    return this;
                }
                else
                {
                    return this;
                }
            }

            public OperationBasic SetDisplay(string text, DisplayType dp = DisplayType.basic)
            {
                display[dp] = text;
                displayOnParse[dp] = text;
                return this;
            }

            public OperationBasic SetDisplayOnly(string text, DisplayType dp = DisplayType.basic)
            {
                display[dp] = text;
                return this;
            }

            public string Display(DisplayType dp=DisplayType.basic){
                if (display.ContainsKey(dp)) { return display[dp]; }
                else if (display.ContainsKey(DisplayType.basic)) { return display[DisplayType.basic]; }
                else { throw new NotImplementedException(); }
            }
        }

    }
}

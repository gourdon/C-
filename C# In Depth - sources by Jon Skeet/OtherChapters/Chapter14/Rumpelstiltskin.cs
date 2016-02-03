using System;
using System.ComponentModel;
using System.Dynamic;
using System.Linq.Expressions;
using System.Reflection;

namespace Chapter14
{
    [Description("Listing 14.35 and 36")]
    public sealed class Rumpelstiltskin : IDynamicMetaObjectProvider
    {
        private readonly string name;
        public Rumpelstiltskin(string name)
        {
            this.name = name;
        }

        public DynamicMetaObject GetMetaObject(Expression expression)
        {
            return new MetaRumpelstiltskin(expression, this);
        }

        private object RespondToWrongGuess(string guess)
        {
            Console.WriteLine("No, I'm not {0}! (I'm {1}.)",
                guess, name);
            return false;
        }

        private object RespondToRightGuess()
        {
            Console.WriteLine("Curses! Foiled again!");
            return true;
        }

        private class MetaRumpelstiltskin : DynamicMetaObject
        {
            private static readonly MethodInfo RightGuessMethod =
                typeof(Rumpelstiltskin).GetMethod("RespondToRightGuess",
                BindingFlags.Instance | BindingFlags.NonPublic);

            private static readonly MethodInfo WrongGuessMethod =
                typeof(Rumpelstiltskin).GetMethod("RespondToWrongGuess",
                BindingFlags.Instance | BindingFlags.NonPublic);

            internal MetaRumpelstiltskin
                (Expression expression, Rumpelstiltskin creator)
                : base(expression, BindingRestrictions.Empty, creator)
            {}

            public override DynamicMetaObject BindInvokeMember
                (InvokeMemberBinder binder, DynamicMetaObject[] args)
            {
                Rumpelstiltskin targetObject = (Rumpelstiltskin)base.Value;
                Expression self = Expression.Convert(base.Expression,
                    typeof(Rumpelstiltskin));

                Expression targetBehavior;
                if (binder.Name == targetObject.name)
                {
                    targetBehavior = Expression.Call(self, RightGuessMethod);
                }
                else
                {
                    targetBehavior = Expression.Call(self, WrongGuessMethod,
                        Expression.Constant(binder.Name));
                }

                var restrictions = BindingRestrictions.GetInstanceRestriction
                    (self, targetObject);
                return new DynamicMetaObject(targetBehavior, restrictions);
            }
        }
    }
}

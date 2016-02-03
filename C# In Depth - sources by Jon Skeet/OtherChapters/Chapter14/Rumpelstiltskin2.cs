using System;
using System.Dynamic;
using System.Linq.Expressions;
using System.Reflection;

namespace Chapter14
{
    /// <summary>
    /// Alternative implementation of Rumpelstiltskin which uses a type restriction instead
    /// of an instance restriction.
    /// </summary>
    public sealed class Rumpelstiltskin2 : IDynamicMetaObjectProvider
    {
        private readonly string name;
        public Rumpelstiltskin2(string name)
        {
            this.name = name;
        }

        public DynamicMetaObject GetMetaObject(Expression expression)
        {
            return new MetaRumpelstiltskin2(expression, this);
        }

        private object RespondToGuess(string guess)
        {
            if (guess == name)
            {
                Console.WriteLine("Curses! Foiled again!");
                return true;
            }
            else
            {
                Console.WriteLine("No, I'm not {0}! (I'm {1}.)",
                    guess, name);
                return false;
            }
        }

        private class MetaRumpelstiltskin2 : DynamicMetaObject
        {
            private static readonly MethodInfo GuessMethod =
                typeof(Rumpelstiltskin).GetMethod("RespondToGuess",
                BindingFlags.Instance | BindingFlags.NonPublic);

            internal MetaRumpelstiltskin2
                (Expression expression, Rumpelstiltskin2 creator)
                : base(expression, BindingRestrictions.Empty, creator)
            {}

            public override DynamicMetaObject BindInvokeMember
                (InvokeMemberBinder binder, DynamicMetaObject[] args)
            {
                Expression self = Expression.Convert(base.Expression,
                    typeof(Rumpelstiltskin2));
                Expression targetBehavior = Expression.Call(self, GuessMethod,
                        Expression.Constant(binder.Name));

                var restrictions = BindingRestrictions.GetTypeRestriction(Expression, typeof(Rumpelstiltskin2));
                return new DynamicMetaObject(targetBehavior, restrictions);
            }
        }
    }
}

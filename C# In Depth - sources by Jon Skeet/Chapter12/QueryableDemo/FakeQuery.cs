using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;

namespace QueryableDemo
{
    [Description("Listing 12.3")]
    class FakeQuery<T> : IQueryable<T>
    {
        public Expression Expression { get; private set; }
        public IQueryProvider Provider { get; private set; }
        public Type ElementType { get; private set; }

        internal FakeQuery(IQueryProvider provider,
                           Expression expression)
        {
            Expression = expression;
            Provider = provider;
            ElementType = typeof(T);
        }

        internal FakeQuery()
            : this(new FakeQueryProvider(), null)
        {
            Expression = Expression.Constant(this);
        }

        public IEnumerator<T> GetEnumerator()
        {
            Logger.Log(this, Expression);
            return Enumerable.Empty<T>().GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            Logger.Log(this, Expression);
            return Enumerable.Empty<T>().GetEnumerator();
        }

        public override string ToString()
        {
            return "FakeQuery";
        }
    }
}

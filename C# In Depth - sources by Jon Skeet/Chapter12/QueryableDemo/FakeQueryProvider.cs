using System.ComponentModel;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System;

namespace QueryableDemo
{
    [Description("Listing 12.4")]
    class FakeQueryProvider : IQueryProvider
    {
        public IQueryable<T> CreateQuery<T>(Expression expression)
        {
            Logger.Log(this, expression);
            return new FakeQuery<T>(this, expression);
        }

        public IQueryable CreateQuery(Expression expression)
        {
            Logger.Log(this, expression);
            Type queryType = typeof(FakeQuery<>).MakeGenericType(expression.Type);
            object[] constructorArgs = new object[] { this, expression };
            return (IQueryable)Activator.CreateInstance(queryType, constructorArgs);
        }

        public T Execute<T>(Expression expression)
        {
            Logger.Log(this, expression);
            return default(T);
        }

        public object Execute(Expression expression)
        {
            Logger.Log(this, expression);
            return null;
        }
    }
}

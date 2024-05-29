using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;

namespace me.fengyj.CommonLib.Utils.Expressions {
    public static class ExpressionUtils {

        private static readonly ConcurrentDictionary<Type, ConcurrentDictionary<string, LambdaExpression>> propertyExps = new();

        /// <summary>
        /// return a string represent the path of a property chain
        /// </summary>
        /// <typeparam name="TInput"></typeparam>
        /// <typeparam name="TResult"></typeparam>
        /// <param name="propertyExp"></param>
        /// <returns></returns>
        /// <remarks>not supports indexed property</remarks>
        /// <example>
        /// <![CDATA[
        /// GetPropertyPath<Post, string>(post => post.Blog.Author.Name) 
        /// >> Blog.Author.Name
        /// ]]>
        /// </example>
        public static string GetPropertyPath<TInput, TResult>(Expression<Func<TInput, TResult>> propertyExp) {

            var exp = propertyExp.Body;
            var propNames = new List<string>();
            while (exp is MemberExpression memberExp) {
                propNames.Add(memberExp.Member.Name);
                exp = memberExp.Expression;
            }
            return string.Join(".", propNames.Reverse<string>().ToArray());
        }

        /// <summary>
        /// return a lambda expression base on the obj type and property path
        /// </summary>
        /// <typeparam name="TObj"></typeparam>
        /// <param name="propertyPath"></param>
        /// <returns></returns>
        /// <remarks>not support indexed property</remarks>
        /// <example>
        /// <![CDATA[
        /// GetPropExp<Post>("Blog.Author.Name")
        /// >> post => post.Blog.Author.Name
        /// ]]>
        /// </example>
        public static LambdaExpression GetPropertyExp<TObj>(string propertyPath) {
            return GetPropertyExp(typeof(TObj), propertyPath);
        }

        /// <summary>
        /// return a lambda expression base on the obj type and property path
        /// </summary>
        /// <param name="type"></param>
        /// <param name="propertyPath"></param>
        /// <returns></returns>
        /// <remarks>not support indexed property</remarks>
        /// <example>
        /// <![CDATA[
        /// GetPropExp<Post>("Blog.Author.Name")
        /// >> post => post.Blog.Author.Name
        /// ]]>
        /// </example>
        public static LambdaExpression GetPropertyExp(Type type, string propertyPath) {

            var propDict = propertyExps.GetOrAdd(type, t => new ConcurrentDictionary<string, LambdaExpression>());
            var lambdaExp = propDict.GetOrAdd(propertyPath, path => {

                var paramExp = Expression.Parameter(type);
                var props = path.Split('.');

                Expression propExp = paramExp;
                var objType = type;
                foreach (var p in props) {
                    var propInfo = objType.GetProperty(p)
                        ?? throw new ArgumentException($"Cannot find the property {p} in {objType.Name}.", nameof(propertyPath));
                    propExp = Expression.Property(propExp, propInfo);
                    objType = propInfo.PropertyType;
                }
                return Expression.Lambda(propExp, paramExp);
            });

            // for outside, we don't know how the expression will be used, 
            // so for safety, return a brand new object
            return ClonePropertyExp(lambdaExp);
        }

        private static LambdaExpression ClonePropertyExp(LambdaExpression lambdaExp) {

            var stack = new Stack<PropertyInfo>();
            var exp = lambdaExp.Body;
            while (exp is MemberExpression mExp && mExp.Member is PropertyInfo pInfo) {
                exp = mExp.Expression;
                stack.Push(pInfo);
            }
            var paramExp = Expression.Parameter(lambdaExp.Parameters[0].Type);
            exp = paramExp;
            while (stack.Count > 0) {
                var memberInfo = stack.Pop();
                exp = Expression.Property(exp, memberInfo);
            }
            return Expression.Lambda(exp, paramExp);
        }

        /// <summary>
        /// return a lambda expression base on the obj type and property path
        /// </summary>
        /// <typeparam name="TObj"></typeparam>
        /// <typeparam name="TProperty"></typeparam>
        /// <param name="propertyPath"></param>
        /// <returns></returns>
        /// <remarks>not supports indexed property</remarks>
        /// <example>
        /// <![CDATA[
        /// GetPropExp<Post>("Blog.Author.Name")
        /// >> post => post.Blog.Author.Name
        /// ]]>
        /// </example>
        public static Expression<Func<TObj, TProperty>> GetPropertyExp<TObj, TProperty>(string propertyPath) {
            if (GetPropertyExp(typeof(TObj), propertyPath) is Expression<Func<TObj, TProperty>> exp)
                return exp;
            throw new ArgumentException($"The property {propertyPath} doesn't match the type {typeof(TProperty).Name}.", nameof(propertyPath));
        }

        /// <summary>
        /// convert a string to a ConstantExpression by specified type
        /// </summary>
        /// <param name="val"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static ConstantExpression GetConstantExpFromString(string val, Type type) {
            return ConstantExpHelper.GetConstantExp(val, type);
        }

        /// <summary>
        /// it provides the converters from a string to object for common struct types, 
        /// and you can call AddParser function to add converts for other types
        /// </summary>
        public static class ConstantExpHelper {

            private static readonly ConcurrentDictionary<Type, Func<string, ConstantExpression>> dict = new();
            private static readonly ConcurrentDictionary<Type, Func<string[], ConstantExpression>> arrDict = new();

            static ConstantExpHelper() {

                dict.TryAdd(typeof(bool), GetExp(bool.Parse));
                dict.TryAdd(typeof(bool?), GetExpForNullable(bool.Parse));
                dict.TryAdd(typeof(byte), GetExp(byte.Parse));
                dict.TryAdd(typeof(byte?), GetExpForNullable(byte.Parse));
                dict.TryAdd(typeof(DateTime), GetExp(DateTime.Parse));
                dict.TryAdd(typeof(DateTime?), GetExpForNullable(DateTime.Parse));
                dict.TryAdd(typeof(DateTimeOffset), GetExp(DateTimeOffset.Parse));
                dict.TryAdd(typeof(DateTimeOffset?), GetExpForNullable(DateTimeOffset.Parse));
                dict.TryAdd(typeof(decimal), GetExp(decimal.Parse));
                dict.TryAdd(typeof(decimal?), GetExpForNullable(decimal.Parse));
                dict.TryAdd(typeof(double), GetExp(double.Parse));
                dict.TryAdd(typeof(double?), GetExpForNullable(double.Parse));
                dict.TryAdd(typeof(Guid), GetExp(Guid.Parse));
                dict.TryAdd(typeof(Guid?), GetExpForNullable(Guid.Parse));
                dict.TryAdd(typeof(short), GetExp(short.Parse));
                dict.TryAdd(typeof(short?), GetExpForNullable(short.Parse));
                dict.TryAdd(typeof(int), GetExp(int.Parse));
                dict.TryAdd(typeof(int?), GetExpForNullable(int.Parse));
                dict.TryAdd(typeof(long), GetExp(long.Parse));
                dict.TryAdd(typeof(long?), GetExpForNullable(long.Parse));
                dict.TryAdd(typeof(sbyte), GetExp(sbyte.Parse));
                dict.TryAdd(typeof(sbyte?), GetExpForNullable(sbyte.Parse));
                dict.TryAdd(typeof(float), GetExp(float.Parse));
                dict.TryAdd(typeof(float?), GetExpForNullable(float.Parse));
                dict.TryAdd(typeof(TimeSpan), GetExp(TimeSpan.Parse));
                dict.TryAdd(typeof(TimeSpan?), GetExpForNullable(TimeSpan.Parse));
                dict.TryAdd(typeof(string), GetExpForStringType());
                dict.TryAdd(typeof(char[]), GetExpForArrayOfCharType());

                arrDict.TryAdd(typeof(bool), GetArrayExp(bool.Parse));
                arrDict.TryAdd(typeof(byte), GetArrayExp(byte.Parse));
                arrDict.TryAdd(typeof(DateTime), GetArrayExp(DateTime.Parse));
                arrDict.TryAdd(typeof(DateTimeOffset), GetArrayExp(DateTimeOffset.Parse));
                arrDict.TryAdd(typeof(decimal), GetArrayExp(decimal.Parse));
                arrDict.TryAdd(typeof(double), GetArrayExp(double.Parse));
                arrDict.TryAdd(typeof(Guid), GetArrayExp(Guid.Parse));
                arrDict.TryAdd(typeof(short), GetArrayExp(short.Parse));
                arrDict.TryAdd(typeof(int), GetArrayExp(int.Parse));
                arrDict.TryAdd(typeof(long), GetArrayExp(long.Parse));
                arrDict.TryAdd(typeof(sbyte), GetArrayExp(sbyte.Parse));
                arrDict.TryAdd(typeof(float), GetArrayExp(float.Parse));
                arrDict.TryAdd(typeof(TimeSpan), GetArrayExp(TimeSpan.Parse));
                arrDict.TryAdd(typeof(string), GetArrayExpForStringType());
            }

            internal static ConstantExpression GetConstantExp(string val, Type type) {
                if (type.IsEnum)
                    return GetExpForEnumType(type, val);
                else if (dict.TryGetValue(type, out var f))
                    return f(val);
                else
                    throw new ArgumentException($"Cannot convert val to {type.Name}.", nameof(type));
            }

            internal static ConstantExpression GetConstantExp(string[] val, Type type) {
                if (type.IsEnum)
                    return GetExpForEnumType(type, val);
                else if (arrDict.TryGetValue(type, out var f))
                    return f(val);
                else
                    throw new ArgumentException($"Cannot convert val to {type.Name}.", nameof(type));
            }

            /// <summary>
            /// add a parser function for specified type to convert the string to the object
            /// </summary>
            /// <typeparam name="T"></typeparam>
            /// <param name="parser"></param>
            public static void AddParser<T>(Func<string, T> parser) {

                var type = typeof(T);
                dict.GetOrAdd(type, t => val => Expression.Constant(parser(val), t));
                arrDict.GetOrAdd(type, t => val => Expression.Constant(val.Select(item => parser(item)).ToArray(), t.MakeArrayType()));
            }

            private static ConstantExpression GetExpForEnumType(Type enumType, string val) {
                var obj = Enum.Parse(enumType, val, true);
                return Expression.Constant(obj, enumType);
            }

            private static ConstantExpression GetExpForEnumType(Type enumType, string[] vals) {
                var objs = vals.Select(v => Enum.Parse(enumType, v, true)).ToArray();
                return Expression.Constant(objs, enumType.MakeArrayType());
            }

            private static Func<string, ConstantExpression> GetExp<T>(Func<string, T> func) where T : struct {
                return val => Expression.Constant(func(val), typeof(T));
            }

            private static Func<string, ConstantExpression> GetExpForNullable<T>(Func<string, T> func) where T : struct {
                return val => {
                    if (string.IsNullOrEmpty(val)) return Expression.Constant(null, typeof(T?));
                    else return Expression.Constant(func(val), typeof(T?));
                };
            }

            private static Func<string, ConstantExpression> GetExpForStringType() {
                return val => Expression.Constant(val, typeof(string));
            }

            private static Func<string, ConstantExpression> GetExpForArrayOfCharType() {
                return val => Expression.Constant(val?.ToCharArray(), typeof(char[]));
            }

            private static Func<string[], ConstantExpression> GetArrayExp<T>(Func<string, T> func) where T : struct {
                return val => Expression.Constant(val.Select(item => func(item)).ToArray(), typeof(T[]));
            }

            private static Func<string[], ConstantExpression> GetArrayExpForStringType() {
                return val => Expression.Constant(val, typeof(string[]));
            }
        }

        public class ParameterUpdateVisitor(ParameterExpression oldParameter, ParameterExpression newParameter) : ExpressionVisitor {
            protected override Expression VisitParameter(ParameterExpression node) {

                if (ReferenceEquals(node, oldParameter))
                    return newParameter;

                return base.VisitParameter(node);
            }
        }

        public class ExpressionReplaceVisitor(Expression oldExp, Expression newExp) : ExpressionVisitor {
            public override Expression? Visit(Expression? node) {

                if (ReferenceEquals(node, oldExp))
                    return newExp;
                return base.Visit(node);
            }
        }
    }
}

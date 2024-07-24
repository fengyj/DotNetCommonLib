using System.Data;
using System.Data.Common;

namespace me.fengyj.CommonLib.Data {
    public class DbCommandHelper<T> {

        private readonly Dictionary<string, Func<T, object?>> valueGetters = [];
        private readonly Dictionary<string, Action<T, object?>> valueSetters = [];

        public DbCommandHelper(DbCommand cmd, DbTransaction? trans = null) {

            this.Command = cmd;
            if (trans != null) this.Command.Transaction = trans;
        }

        public DbCommandHelper(DbConnection conn, string sql) : this(conn.CreateCommand()) {

            this.Command.CommandText = sql;
        }

        public DbCommandHelper(DbTransaction trans, string sql) {

            if (trans.Connection == null) throw new ArgumentException("Transaction's connection is null.", nameof(trans));
            this.Command = trans.Connection.CreateCommand();
            this.Command.CommandText = sql;
            this.Command.Transaction = trans;
        }

        public DbCommand Command { get; private set; }

        public DbCommandHelper<T> AddParameter(
            string paramName,
            DbType dbType,
            Func<T, object?> valueGetter,
            int? size = null,
            byte? scale = null,
            byte? precision = null) {

            return this.AddParameter(paramName, dbType, valueGetter, null, size: size, scale: scale, precision: precision);
        }

        public DbCommandHelper<T> AddParameter(
            string paramName,
            DbType dbType,
            Action<T, object?> valueSetter,
            int? size = null,
            byte? scale = null,
            byte? precision = null) {
            return this.AddParameter(paramName, dbType, null, valueSetter, size: size, scale: scale, precision: precision);
        }

        public DbCommandHelper<T> AddParameter(
            string paramName,
            DbType dbType,
            Func<T, object?>? valueGetter,
            Action<T, object?>? valueSetter,
            int? size = null,
            byte? scale = null,
            byte? precision = null,
            ParameterDirection? direction = null) {

            var param = this.Command.CreateParameter();

            param.ParameterName = paramName;
            param.DbType = dbType;

            if (direction == null) {
                if (valueGetter != null && valueSetter != null)
                    param.Direction = ParameterDirection.InputOutput;
                else if (valueSetter == null)
                    param.Direction = ParameterDirection.Input;
                else
                    param.Direction = ParameterDirection.Output;
            }
            else {
                param.Direction = direction.Value;
            }
            if (size.HasValue) param.Size = size.Value;
            if (precision.HasValue) param.Precision = precision.Value;
            if (scale.HasValue) param.Scale = scale.Value;

            this.Command.Parameters.Add(param);

            if (valueGetter != null) this.valueGetters.Add(paramName, valueGetter);
            if (valueSetter != null) this.valueSetters.Add(paramName, valueSetter);

            return this;
        }

        public int ExecuteNonQuery(T data) {

            foreach (var item in this.valueGetters) {
                var val = item.Value(data);
                this.Command.Parameters[item.Key].Value = val ?? DBNull.Value;
            }
            return this.Command.ExecuteNonQuery();
        }

        public object? ExecuteScalar(T data) {

            foreach (var item in this.valueGetters) {
                var val = item.Value(data);
                this.Command.Parameters[item.Key].Value = val ?? DBNull.Value;
            }
            var result = this.Command.ExecuteScalar();
            if (Convert.IsDBNull(result)) return null;
            else return result;
        }

        public V? ExecuteScalar<V>(T data) {

            var r = this.ExecuteScalar(data);
            if (r == null || r is not V v) return default;
            return v;
        }

        public DbDataReader ExecuteReader(T data) {

            foreach (var item in this.valueGetters) {
                var val = item.Value(data);
                this.Command.Parameters[item.Key].Value = val ?? DBNull.Value;
            }
            return this.Command.ExecuteReader();
        }

        public async Task<int> ExecuteNonQueryAsync(T data) {

            foreach (var item in this.valueGetters) {
                var val = item.Value(data);
                this.Command.Parameters[item.Key].Value = val ?? DBNull.Value;
            }
            return await this.Command.ExecuteNonQueryAsync();
        }

        public async Task<object?> ExecuteScalarAsync(T data) {

            foreach (var item in this.valueGetters) {
                var val = item.Value(data);
                this.Command.Parameters[item.Key].Value = val ?? DBNull.Value;
            }
            var result = await this.Command.ExecuteScalarAsync();
            if (Convert.IsDBNull(result)) return null;
            else return result;
        }

        public async Task<V?> ExecuteScalarAsync<V>(T data) {

            var r = await this.ExecuteScalarAsync(data);
            if (r == null || r is not V v) return default;
            return v;
        }

        public async Task<DbDataReader> ExecuteReaderAsync(T data) {

            foreach (var item in this.valueGetters) {
                var val = item.Value(data);
                this.Command.Parameters[item.Key].Value = val ?? DBNull.Value;
            }
            return await this.Command.ExecuteReaderAsync();
        }
    }

    public class DbCommandHelper {

        public DbCommandHelper(DbCommand cmd, DbTransaction? trans = null) {

            this.Command = cmd;
            if (trans != null) this.Command.Transaction = trans;
        }

        public DbCommandHelper(DbConnection conn, string sql) : this(conn.CreateCommand()) {

            this.Command.CommandText = sql;
        }

        public DbCommandHelper(DbTransaction trans, string sql) {

            if (trans.Connection == null) throw new ArgumentException("Transaction's connection is null.", nameof(trans));
            this.Command = trans.Connection.CreateCommand();
            this.Command.CommandText = sql;
            this.Command.Transaction = trans;
        }

        public DbCommand Command { get; private set; }

        public DbCommandHelper AddParameter(
            string paramName,
            DbType dbType,
            int? size = null,
            byte? scale = null,
            byte? precision = null,
            ParameterDirection direction = ParameterDirection.Output) {
            return this.AddParameter(paramName, null, dbType, size: size, scale: scale, precision: precision, direction: direction);
        }

        public DbCommandHelper AddParameter(
            string paramName,
            object? value,
            DbType dbType,
            int? size = null,
            byte? scale = null,
            byte? precision = null,
            ParameterDirection direction = ParameterDirection.Input) {

            var param = this.Command.CreateParameter();

            param.ParameterName = paramName;
            param.Direction = direction;
            param.DbType = dbType;
            if (size.HasValue) param.Size = size.Value;
            if (precision.HasValue) param.Precision = precision.Value;
            if (scale.HasValue) param.Scale = scale.Value;

            if (direction == ParameterDirection.Input || direction == ParameterDirection.InputOutput)
                param.Value = value;

            this.Command.Parameters.Add(param);

            return this;
        }

        public int ExecuteNonQuery() {

            return this.Command.ExecuteNonQuery();
        }

        public object? ExecuteScalar() {

            var result = this.Command.ExecuteScalar();
            if (Convert.IsDBNull(result)) return null;
            else return result;
        }

        public V? ExecuteScalar<V>() {

            var r = this.ExecuteScalar();
            if (r == null || r is not V v) return default;
            return v;
        }

        public DbDataReader ExecuteReader() {

            return this.Command.ExecuteReader();
        }

        public IEnumerable<V> Fetch<V>() {

            using (var reader = this.ExecuteReader()) {
                yield break;
            }
        }

        public async Task<int> ExecuteNonQueryAsync() {

            return await this.Command.ExecuteNonQueryAsync();
        }

        public async Task<object?> ExecuteScalarAsync() {

            var result = await this.Command.ExecuteScalarAsync();
            if (Convert.IsDBNull(result)) return null;
            else return result;
        }

        public async Task<V?> ExecuteScalarAsync<V>() {

            var r = await this.ExecuteScalarAsync();
            if (r == null || r is not V v) return default;
            return v;
        }

        public async Task<DbDataReader> ExecuteReaderAsync() {

            return await this.Command.ExecuteReaderAsync();
        }

        public async Task<IEnumerable<V>> FetchAsync<V>() {

            using (var reader = await this.ExecuteReaderAsync()) {

                return [];
            }
        }

    }
}

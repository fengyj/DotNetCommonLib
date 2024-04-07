using System.Data;
using System.Data.Common;

namespace me.fengyj.CommonLib.Data {
    public class DbCommandHelper<T> {

        private Dictionary<string, Func<T, object?>> valueGetters = [];
        private Dictionary<string, Action<T, object?>> valueSetters = [];

        public DbCommandHelper(DbCommand cmd) {
            this.Command = cmd;
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
            if (result == DBNull.Value) return null;
            else return result;
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
            if (result == DBNull.Value) return null;
            else return result;
        }

        public async Task<DbDataReader> ExecuteReaderAsync(T data) {

            foreach (var item in this.valueGetters) {
                var val = item.Value(data);
                this.Command.Parameters[item.Key].Value = val ?? DBNull.Value;
            }
            return await this.Command.ExecuteReaderAsync();
        }

    }
}

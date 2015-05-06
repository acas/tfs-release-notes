using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.SQLite;
using System.Data;

namespace ReleaseNotesWeb.SQLite
{
    public class BasicSQLiteDriver
    {
        public enum ColumnDataType
        {
            INTEGER, 
            REAL, 
            TEXT, 
            BLOB, 
            NULL
        }

        public enum TableConstraintType
        {
            PRIMARY, 
            FOREIGN, 
            UNIQUE, 
            CHECK
        }

        public enum ColumnConstraintType
        {
            PRIMARY,
            NOTNULL,
            CHECK,
            DEFAULT,
            UNIQUE,
            FOREIGN
        }

        public class TableConstraint
        {
            private TableConstraintType constraintType;
            private string referentialTable;
            private List<string> appliedColumns = new List<string>();
            private List<string> referentialColumns = new List<string>();
            private string appliedTable;
            private string appliedColumn;
            private string defaultValueOrExpr;

            public string ReferentialTable
            {
                get { return referentialTable; }
                set { referentialTable = value; }
            }

            public List<string> AppliedColumns
            {
                get { return appliedColumns; }
                set { appliedColumns = value; }
            }

            public List<string> ReferentialColumns
            {
                get { return referentialColumns; }
                set { referentialColumns = value; }
            }

            public string DefaultValueOrExpr
            {
                get { return defaultValueOrExpr; }
                set { defaultValueOrExpr = value; }
            }

            public TableConstraintType ConstraintType
            {
                get { return constraintType; }
                set { constraintType = value; }
            }

            public string AppliedTable
            {
                get { return appliedTable; }
                set { appliedTable = value; }
            }

            public string AppliedColumn
            {
                get { return appliedColumn; }
                set { appliedColumn = value; }
            }
        }

        public class ColumnConstraint
        {
            private ColumnConstraintType constraintType;
            private string referentialTable;
            private List<string> referentialColumns = new List<string>();
            private string defaultValueOrExpr;
            private bool autoIncrementColumn = false;

            public bool AutoIncrementColumn
            {
                get { return autoIncrementColumn; }
                set { autoIncrementColumn = value; }
            }

            public List<string> ReferentialColumns
            {
                get { return referentialColumns; }
                set { referentialColumns = value; }
            }

            public string ReferentialTable
            {
                get { return referentialTable; }
                set { referentialTable = value; }
            }

            public ColumnConstraintType ConstraintType
            {
                get { return constraintType; }
                set { constraintType = value; }
            }

            public string DefaultValueOrExpr
            {
                get { return defaultValueOrExpr; }
                set { defaultValueOrExpr = value; }
            }
        }

        public class DriverTable
        {
            private string tableName;
            private List<DriverColumn> columns = new List<DriverColumn>();
            private List<TableConstraint> keys = new List<TableConstraint>();

            public string TableName
            {
                get { return tableName; }
                set { tableName = value; }
            }

            public List<DriverColumn> Columns
            {
                get { return columns; }
                set { columns = value; }
            }

            public List<TableConstraint> Keys
            {
                get { return keys; }
                set { keys = value; }
            }
        }

        public class DriverColumn 
        {
            private string columnName;
            private ColumnDataType columnType;
            private List<ColumnConstraint> constraints = new List<ColumnConstraint>();

            public List<ColumnConstraint> Constraints
            {
                get { return constraints; }
                set { constraints = value; }
            }
            
            public string ColumnName
            {
                get { return columnName; }
                set { columnName = value; }
            }

            public ColumnDataType ColumnType
            {
                get { return columnType; }
                set { columnType = value; }
            }
        }

        private SQLiteConnection connection = null;

        private BasicSQLiteDriver(string connectionString)
        {
            connection = new SQLiteConnection(connectionString);
            connection.Open();
        }

        public static BasicSQLiteDriver CreateDriver(string connectionString) {
            try
            {
                BasicSQLiteDriver d = new BasicSQLiteDriver(connectionString);
                return d;
            }
            catch (SQLiteException)
            {
                throw;
            }
        }

        private void RunNonQuery(string sqlString)
        {
            try
            {
                SQLiteCommand command = new SQLiteCommand(sqlString, this.connection);
                command.ExecuteNonQuery();
            }
            catch (SQLiteException)
            {
                throw;
            }
            catch (NullReferenceException) 
            {
                throw;
            }
        }

        public DataTable RunQuery(string sqlString)
        {
            try
            {
                DataTable dt = new DataTable("Result");
                SQLiteCommand command = new SQLiteCommand(sqlString, this.connection);
                SQLiteDataAdapter da = new SQLiteDataAdapter(command);
                da.Fill(dt);
                return dt;
            }
            catch (SQLiteException)
            {
                throw;
            }
            catch (NullReferenceException)
            {
                throw;
            }
        }

        public bool GetBasicExistsQueryResult(string tableName, string columnName, DbType columnType, object columnValue)
        {
            try
            {
                string resultColumnName = "dataexists";
                string sqlString = "select count(*) as " + resultColumnName + " from " + tableName + " where " + columnName + " = ?;";
                SQLiteCommand command = new SQLiteCommand(sqlString, connection);
                command.Parameters.Add(new SQLiteParameter(columnType, columnValue));
                command.ExecuteNonQuery();
                DataTable dt = new DataTable("Result");
                SQLiteDataAdapter da = new SQLiteDataAdapter(command);
                da.Fill(dt);
                return int.Parse(dt.Rows[0][0].ToString()) > 0;
            }
            catch (ArgumentOutOfRangeException)
            {
                throw;
            }
            catch (SQLiteException)
            {
                throw;
            }
        }

        public void CreateBasicUpdateStatement(string tableName, List<Tuple<string, DbType>> columns, List<object> values, List<Tuple<string, DbType>> filterColumns, List<object> filterValues)
        {
            try
            {
                if (columns.Count() != values.Count())
                    throw new ArgumentOutOfRangeException("Column and value list sizes do not match");
                if (filterColumns.Count() != filterValues.Count())
                    throw new ArgumentOutOfRangeException("Filter column and value list sizes do not match");

                string sqlString = "update " + tableName + " set ";
                for (var i = 0; i < columns.Count; i++)
                {
                    sqlString += columns.ElementAt(i).Item1 + " = ?";
                    if (i != columns.Count - 1)
                    {
                        sqlString += ", ";
                    }
                }
                sqlString += " " + "where" + " ";
                for (var i = 0; i < filterColumns.Count; i++)
                {
                    sqlString += filterColumns.ElementAt(i).Item1 + " = ?";
                    if (i != filterColumns.Count - 1)
                    {
                        sqlString += "and ";
                    }
                }
                sqlString += ";";

                SQLiteCommand command = new SQLiteCommand(sqlString, connection);
                for (var i = 0; i < values.Count; i++)
                {
                    command.Parameters.Add(new SQLiteParameter(columns.ElementAt(i).Item2, values.ElementAt(i)));
                }
                for (var i = 0; i < filterValues.Count; i++)
                {
                    command.Parameters.Add(new SQLiteParameter(filterColumns.ElementAt(i).Item2, filterValues.ElementAt(i)));
                }
                command.ExecuteNonQuery();

            }
            catch (ArgumentOutOfRangeException)
            {
                throw;
            }
            catch (SQLiteException)
            {
                throw;
            }
        }

        public void CreateBasicDeleteStatement(string tableName, List<Tuple<string, DbType>> columns, List<object> values)
        {
            try
            {
                if (columns.Count() != values.Count())
                    throw new ArgumentOutOfRangeException("Column and value list sizes do not match");

                string sqlString = "delete from " + tableName + " where ";
                for (var i = 0; i < columns.Count; i++)
                {
                    sqlString += columns.ElementAt(i).Item1 + " = ?";
                    if (i != columns.Count - 1)
                    {
                        sqlString += "and ";
                    }
                }
                sqlString += ";";

                SQLiteCommand command = new SQLiteCommand(sqlString, connection);
                for (var i = 0; i < values.Count; i++)
                {
                    command.Parameters.Add(new SQLiteParameter(columns.ElementAt(i).Item2, values.ElementAt(i)));
                }
                command.ExecuteNonQuery();

            }
            catch (ArgumentOutOfRangeException)
            {
                throw;
            }
            catch (SQLiteException)
            {
                throw;
            }
        }

        public void CreateBasicInsertStatement(string tableName, List<Tuple<string, DbType>> columns, List<object> values)
        {
            try
            {
                if (columns.Count() != values.Count())
                    throw new ArgumentOutOfRangeException("Column and value list sizes do not match");

                string sqlString = "insert into " + tableName + "(";
                for (var i = 0; i < columns.Count; i++)
                {
                    sqlString += columns.ElementAt(i).Item1;
                    if (i != columns.Count - 1)
                    {
                        sqlString += ", ";
                    }
                }

                sqlString += ") values (";
                for (var i = 0; i < values.Count; i++)
                {
                    sqlString += "?";
                    if (i != values.Count - 1)
                    {
                        sqlString += ", ";
                    }
                }
                sqlString += ");";

                SQLiteCommand command = new SQLiteCommand(sqlString, connection);
                for (var i = 0; i < values.Count; i++)
                {
                    command.Parameters.Add(new SQLiteParameter(columns.ElementAt(i).Item2, values.ElementAt(i)));
                }
                command.ExecuteNonQuery();

            }
            catch (ArgumentOutOfRangeException)
            {
                throw;
            }
            catch (SQLiteException)
            {
                throw;
            }
        }

        public void CreateTable(DriverTable schema)
        {
            string sqlString = "create table " + schema.TableName + "( ";
            for (var i = 0; i < schema.Columns.Count(); i++)
            {
                sqlString += schema.Columns.ElementAt(i).ColumnName + " ";
                switch (schema.Columns.ElementAt(i).ColumnType)
                {
                    case ColumnDataType.INTEGER:
                        sqlString += "integer";
                        break;
                    case ColumnDataType.REAL:
                        sqlString += "real";
                        break;
                    case ColumnDataType.TEXT:
                        sqlString += "text";
                        break;
                    case ColumnDataType.BLOB:
                        sqlString += "blob";
                        break;
                    default:
                        throw new Exception("Create table: Column type invalid.");
                }

                if (schema.Columns.ElementAt(i).Constraints.Count() > 0)
                {
                    sqlString += " ";
                    for (var j = 0; j < schema.Columns.ElementAt(i).Constraints.Count(); j++)
                    {
                        switch (schema.Columns.ElementAt(i).Constraints.ElementAt(j).ConstraintType)
                        {
                            case ColumnConstraintType.PRIMARY:
                                sqlString += "primary key";
                                if (schema.Columns.ElementAt(i).Constraints.ElementAt(j).AutoIncrementColumn == true)
                                {
                                    sqlString += " " + "autoincrement";
                                }
                                break;
                            case ColumnConstraintType.NOTNULL:
                                sqlString += "not null";
                                break;
                            case ColumnConstraintType.UNIQUE:
                                sqlString += "unique";
                                break;
                            case ColumnConstraintType.CHECK:
                                sqlString += "check (" + schema.Columns.ElementAt(i).Constraints.ElementAt(j).DefaultValueOrExpr + ")";
                                break;
                            case ColumnConstraintType.DEFAULT:
                                sqlString += "default (" + schema.Columns.ElementAt(i).Constraints.ElementAt(j).DefaultValueOrExpr + ")";
                                break;
                            case ColumnConstraintType.FOREIGN:
                                sqlString += "references " + schema.Columns.ElementAt(i).Constraints.ElementAt(j).ReferentialTable + "(";
                                for (var k = 0; k < schema.Columns.ElementAt(i).Constraints.ElementAt(j).ReferentialColumns.Count(); k++)
                                {
                                    sqlString += schema.Columns.ElementAt(i).Constraints.ElementAt(j).ReferentialColumns.ElementAt(k);
                                    if (k != schema.Columns.ElementAt(i).Constraints.ElementAt(j).ReferentialColumns.Count() - 1)
                                    {
                                        sqlString += ", ";
                                    }
                                }
                                sqlString += ")";
                                break;
                            default:
                                throw new Exception("Create table:  Column constraint type invalid.");
                        }
                        if (j != schema.Columns.ElementAt(i).Constraints.Count() - 1)
                        {
                            sqlString += " ";
                        }
                    }
                }

                if (i != schema.Columns.Count() - 1)
                {
                    sqlString += ", ";
                }
            }

            if (schema.Keys.Count() > 0)
            {
                for (var i = 0; i < schema.Keys.Count(); i++)
                {
                    switch (schema.Keys.ElementAt(i).ConstraintType)
                    {
                        case TableConstraintType.CHECK:
                            sqlString += "check (" + schema.Keys.ElementAt(i).DefaultValueOrExpr + ")"; 
                            break;
                        case TableConstraintType.UNIQUE:
                            sqlString += "unique";
                            break;
                        case TableConstraintType.FOREIGN:
                            sqlString += "foreign key (";
                            for (var k = 0; k < schema.Keys.ElementAt(i).AppliedColumns.Count(); k++)
                            {
                                sqlString += schema.Keys.ElementAt(i).AppliedColumns.ElementAt(k);
                                if (k != schema.Keys.ElementAt(i).AppliedColumns.Count() - 1)
                                {
                                    sqlString += ", ";
                                }
                            }
                            sqlString += ")";
                            sqlString += "references " + schema.Keys.ElementAt(i).ReferentialTable + "(";
                                for (var k = 0; k < schema.Keys.ElementAt(i).ReferentialColumns.Count(); k++)
                                {
                                    sqlString += schema.Keys.ElementAt(i).ReferentialColumns.ElementAt(k);
                                    if (k != schema.Keys.ElementAt(i).ReferentialColumns.Count() - 1)
                                    {
                                        sqlString += ", ";
                                    }
                                }
                                sqlString += ")";
                            break;
                        case TableConstraintType.PRIMARY:
                            break;
                    }

                    if (i != schema.Keys.Count() - 1)
                    {
                        sqlString += " ";
                    }
                }
            }
            sqlString += " );";

            // now run
            RunNonQuery(sqlString);
        }
    }
}
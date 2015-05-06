using Newtonsoft.Json.Linq;
using ReleaseNotesWeb.SQLite;
using ReleaseNotesWeb.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ReleaseNotesWeb.Controllers
{
    [RoutePrefix("api/Configuration")]
    public class ConfigurationController : ApiController
    {
        internal static BasicSQLiteDriver d = BasicSQLiteDriver.CreateDriver("Data Source=" + ReleaseNotesWeb.WebApiApplication.configDbPath + ";Version=3;");
        internal static List<Tuple<string, DbType>> columns = new List<Tuple<string, DbType>>
        {
            new Tuple<string, DbType>("generator", DbType.String),
            new Tuple<string, DbType>("teamProjectPath", DbType.String),
            new Tuple<string, DbType>("projectName", DbType.String),
            new Tuple<string, DbType>("projectSubpath", DbType.String),
            new Tuple<string, DbType>("iteration", DbType.String),
            new Tuple<string, DbType>("database", DbType.String),
            new Tuple<string, DbType>("databaseServer", DbType.String),
            new Tuple<string, DbType>("webServer", DbType.String),
            new Tuple<string, DbType>("webLocation", DbType.String)
        };

        [Route("Load")]
        [HttpGet]
        public DataTable LoadConfigurations()
        {
            return d.RunQuery("select * from configurations");
        }

        [Route("Save")]
        [HttpPost]
        public void SaveConfiguration([FromBody] JObject fields)
        {
            List<object> values = new List<object> 
            {
                fields.GetValueOrDefault<string>("generator"),
                fields.GetValueOrDefault<string>("teamProjectPath"),
                fields.GetValueOrDefault<string>("projectName"),
                fields.GetValueOrDefault<string>("projectSubpath"),
                fields.GetValueOrDefault<string>("iteration"),
                fields.GetValueOrDefault<string>("database"),
                fields.GetValueOrDefault<string>("databaseServer"),
                fields.GetValueOrDefault<string>("webServer"),
                fields.GetValueOrDefault<string>("webLocation")
            };

            bool exists = d.GetBasicExistsQueryResult("configurations", "configurationName", DbType.String, fields.GetValueOrDefault<string>("configurationName"));
            if (exists)
            {
                d.CreateBasicUpdateStatement("configurations", columns, values,
                new List<Tuple<string, DbType>>
                {
                    new Tuple<string, DbType>("configurationName", DbType.String)
                },
                new List<object>
                {
                    fields.GetValueOrDefault<string>("configurationName")
                });
            }
            else
            {
                List<Tuple<string, DbType>> insertColumns = new List<Tuple<string, DbType>>();
                insertColumns.AddRange(columns);
                insertColumns.Add(new Tuple<string, DbType>("configurationName", DbType.String));
                values.Add(fields.GetValueOrDefault<string>("configurationName"));
                d.CreateBasicInsertStatement("configurations", insertColumns, values);
            }
        }

        [Route("Delete")]
        [HttpPost]
        public void DeleteConfiguration([FromBody] JObject fields)
        {
            List<Tuple<string, DbType>> deleteColumns = new List<Tuple<string, DbType>> {
                new Tuple<string, DbType>("configurationName", DbType.String)
            };
            List<object> values = new List<object> 
            {
                fields.GetValueOrDefault<string>("configurationName")
            };
            bool exists = d.GetBasicExistsQueryResult("configurations", "configurationName", DbType.String, (string)fields.GetValueOrDefault<string>("configurationName"));
            if (exists)
            {
                d.CreateBasicDeleteStatement("configurations", deleteColumns, values);
            }
        }
    }
}

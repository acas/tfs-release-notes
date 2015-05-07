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
    [RoutePrefix("api/Preset")]
    public class PresetController : ApiController
    {
        internal static BasicSQLiteDriver d = BasicSQLiteDriver.CreateDriver("Data Source=" + ReleaseNotesWeb.WebApiApplication.presetsDbPath + ";Version=3;");
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
        public DataTable LoadPresets()
        {
            return d.RunQuery("select * from presets");
        }

        [Route("Save")]
        [HttpPost]
        public void SavePreset([FromBody] JObject fields)
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

            bool exists = d.GetBasicExistsQueryResult("presets", "presetName", DbType.String, fields.GetValueOrDefault<string>("presetName"));
            if (exists)
            {
                d.CreateBasicUpdateStatement("presets", columns, values,
                new List<Tuple<string, DbType>>
                {
                    new Tuple<string, DbType>("presetName", DbType.String)
                },
                new List<object>
                {
                    fields.GetValueOrDefault<string>("presetName")
                });
            }
            else
            {
                List<Tuple<string, DbType>> insertColumns = new List<Tuple<string, DbType>>();
                insertColumns.AddRange(columns);
                insertColumns.Add(new Tuple<string, DbType>("presetName", DbType.String));
                values.Add(fields.GetValueOrDefault<string>("presetName"));
                d.CreateBasicInsertStatement("presets", insertColumns, values);
            }
        }

        [Route("Delete")]
        [HttpPost]
        public void DeletePreset([FromBody] JObject fields)
        {
            List<Tuple<string, DbType>> deleteColumns = new List<Tuple<string, DbType>> {
                new Tuple<string, DbType>("presetName", DbType.String)
            };
            List<object> values = new List<object> 
            {
                fields.GetValueOrDefault<string>("presetName")
            };
            bool exists = d.GetBasicExistsQueryResult("presets", "presetName", DbType.String, (string)fields.GetValueOrDefault<string>("presetName"));
            if (exists)
            {
                d.CreateBasicDeleteStatement("presets", deleteColumns, values);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using System.Data.SQLite;
using System.IO;
using ReleaseNotesWeb.SQLite;

namespace ReleaseNotesWeb
{
    public class WebApiApplication : System.Web.HttpApplication
    {
        public static string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ReleaseNotes\\";
        public static string dbName = "presets.sqlite";
        public static string presetsDbPath = appDataPath + dbName;

        public static ReleaseNotesWeb.SQLite.BasicSQLiteDriver.DriverTable presetTable = new ReleaseNotesWeb.SQLite.BasicSQLiteDriver.DriverTable
        {
            TableName = "presets",
            Columns = new List<BasicSQLiteDriver.DriverColumn> {
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "id",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.INTEGER,
                            Constraints = new List<BasicSQLiteDriver.ColumnConstraint> {
                                new BasicSQLiteDriver.ColumnConstraint {
                                    ConstraintType = BasicSQLiteDriver.ColumnConstraintType.PRIMARY,
                                    AutoIncrementColumn = true
                                },
                                new BasicSQLiteDriver.ColumnConstraint {
                                    ConstraintType = BasicSQLiteDriver.ColumnConstraintType.NOTNULL
                                }
                            }
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "presetName",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT,
                            Constraints = new List<BasicSQLiteDriver.ColumnConstraint> {
                                new BasicSQLiteDriver.ColumnConstraint {
                                    ConstraintType = BasicSQLiteDriver.ColumnConstraintType.UNIQUE
                                }
                            }
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "generator",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "teamProjectPath",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "projectName",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "projectSubpath",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "iteration",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "database",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "databaseServer",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "webServer",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        },
                        new BasicSQLiteDriver.DriverColumn {
                            ColumnName = "webLocation",
                            ColumnType = BasicSQLiteDriver.ColumnDataType.TEXT
                        }
                    }
        };

        protected void Application_Start()
        {
            // initialize WebAPI
            AreaRegistration.RegisterAllAreas();
            GlobalConfiguration.Configure(WebApiConfig.Register);
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);

            // initialize SQLite
            // get app data folder
            if (!Directory.Exists(appDataPath))
            {
                Directory.CreateDirectory(appDataPath);
            }

            // create presets file if it didn't exist before
            if (!File.Exists(presetsDbPath))
            {
                SQLiteConnection.CreateFile(presetsDbPath);
                BasicSQLiteDriver d = BasicSQLiteDriver.CreateDriver("Data Source=" + presetsDbPath + ";Version=3;");
                d.CreateTable(presetTable);
            }
        }
    }
}

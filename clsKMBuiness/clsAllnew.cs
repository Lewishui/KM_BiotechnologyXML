using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using clsdatabaseinfo;
using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;

namespace clsKMBuiness
{
    public class clsAllnew
    {

        string connectionString = "mongodb://127.0.0.1";
        string DB_NAME = "KM_BiotechnologyXML";
        public clsAllnew()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "System\\IP.txt";

            string[] fileText = File.ReadAllLines(path);
            connectionString = "mongodb://" + fileText[0];


        }
        public void SPInputclaimreport_Server(List<xmlDataSources> AddMAPResult)
        {

            MongoServer server = MongoServer.Create(connectionString);
            MongoDatabase db1 = server.GetDatabase(DB_NAME);
            MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlData");
            MongoCollection<BsonDocument> employees1 = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlData");

            // collection1.RemoveAll();
            if (AddMAPResult == null)
            {
                MessageBox.Show("No Data  input Sever", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            foreach (xmlDataSources item in AddMAPResult)
            {
                var dd = Query.And(Query.EQ("Rack_ID", item.Rack_ID), Query.EQ("Tube_ID", item.Tube_ID), Query.EQ("Hole_ID", item.Hole_ID));//同时满足多个条件


                collection1.Remove(dd);

                MongoDatabase db = server.GetDatabase(DB_NAME);
                MongoCollection collection = db.GetCollection("KM_BiotechnologyXML_XmlData");
                BsonDocument fruit_1 = new BsonDocument
                 { 
                 { "Rack_ID", item.Rack_ID },
                 { "Hole_ID", item.Hole_ID },
                 { "Input_Date", item.Input_Date}, 
                 { "Tube_ID", item.Tube_ID} ,
                 { "FileName", item.FileName} ,
                 { "OrderName", item.OrderName} ,
                 { "Comments", item.Comments} 
              
                 };
                collection.Insert(fruit_1);
            }
        }
        public void status_Server(List<string> AddMAPResult)
        {

            MongoServer server = MongoServer.Create(connectionString);
            MongoDatabase db1 = server.GetDatabase(DB_NAME);
            MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlStaus");
            MongoCollection<BsonDocument> employees1 = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlStaus");

            collection1.RemoveAll();
            if (AddMAPResult == null)
            {
                MessageBox.Show("No Data  input Sever", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            for (int i = 0; i < AddMAPResult.Count; i++)
            {
                //var dd = Query.And(Query.EQ("Rack_ID", AddMAPResult[i]));//同时满足多个条件
                //collection1.Remove(dd);
                MongoDatabase db = server.GetDatabase(DB_NAME);
                MongoCollection collection = db.GetCollection("KM_BiotechnologyXML_XmlStaus");
                BsonDocument fruit_1 = new BsonDocument
                 { 
                 { "Rack_ID", AddMAPResult[i] },              
                 { "Input_Date", DateTime.Now.ToString("yyyyMMdd")}
              
                 };
                collection.Insert(fruit_1);
            }
        }
        public List<xmlDataSources> findALLStatus_Server()
        {


            try
            {
                List<xmlDataSources> ClaimReport_Server = new List<xmlDataSources>();

                MongoServer server = MongoServer.Create(connectionString);
                MongoDatabase db1 = server.GetDatabase(DB_NAME);
                MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlStaus");
                MongoCollection<BsonDocument> employees = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlStaus");

                foreach (BsonDocument emp in employees.FindAll())
                {
                    xmlDataSources item = new xmlDataSources();

                    #region 数据
                    if (emp.Contains("_id"))
                        item.Id = (emp["_id"].ToString());
                    if (emp.Contains("Rack_ID"))
                        item.Rack_ID = (emp["Rack_ID"].ToString());
                  
                    if (emp.Contains("Input_Date"))
                        item.Input_Date = (emp["Input_Date"].AsString);

                    #endregion

                    ClaimReport_Server.Add(item);
                }
                return ClaimReport_Server;

            }
            catch (Exception ex)
            {
                MessageBox.Show("信息读取异常， 请检查网络或IP是否正确配置" + ex);
                return null;

                throw ex;
            }
        }
        public List<xmlDataSources> findALLXML_Server()
        {


            try
            {
                List<xmlDataSources> ClaimReport_Server = new List<xmlDataSources>();

                MongoServer server = MongoServer.Create(connectionString);
                MongoDatabase db1 = server.GetDatabase(DB_NAME);
                MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlData");
                MongoCollection<BsonDocument> employees = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlData");

                foreach (BsonDocument emp in employees.FindAll())
                {
                    xmlDataSources item = new xmlDataSources();

                    #region 数据
                    if (emp.Contains("_id"))
                        item.Id = (emp["_id"].ToString());
                    if (emp.Contains("Rack_ID"))
                        item.Rack_ID = (emp["Rack_ID"].ToString());
                    if (emp.Contains("Hole_ID"))
                        item.Hole_ID = (emp["Hole_ID"].ToString());
                    if (emp.Contains("Tube_ID"))
                        item.Tube_ID = (emp["Tube_ID"].ToString());
                    if (emp.Contains("FileName"))
                        item.FileName = (emp["FileName"].AsString);
                    if (emp.Contains("Comments"))
                        item.Comments = (emp["Comments"].AsString);
                    if (emp.Contains("OrderName"))
                        item.OrderName = (emp["OrderName"].AsString);
                    if (emp.Contains("Input_Date"))
                        item.Input_Date = (emp["Input_Date"].AsString);

                    #endregion

                    ClaimReport_Server.Add(item);
                }
                return ClaimReport_Server;

            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);
                return null;

                throw ex;
            }
        }
        public List<xmlDataSources> findXML_Server(string kettext, string start_time, string end_time, string Rack_ID)
        {

            #region Read  database info server
            try
            {
                List<xmlDataSources> ClaimReport_Server = new List<xmlDataSources>();

                MongoServer server = MongoServer.Create(connectionString);
                MongoDatabase db1 = server.GetDatabase(DB_NAME);
                MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlData");
                MongoCollection<BsonDocument> employees = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlData");

                var query = new QueryDocument("Comments", kettext);
                //    var dd = Query.And(Query.EQ("jigoudaima", jigoudaima), Query.EQ("fapiaoleixing", fapiaoleixing));//同时满足多个条件
                if (kettext != "")
                {
                    foreach (BsonDocument emp in employees.Find(query))
                    {
                        xmlDataSources item = new xmlDataSources();

                        #region 数据
                        if (emp.Contains("_id"))
                            item.Id = (emp["_id"].ToString());
                        if (emp.Contains("Rack_ID"))
                            item.Rack_ID = (emp["Rack_ID"].ToString());
                        if (emp.Contains("Hole_ID"))
                            item.Hole_ID = (emp["Hole_ID"].ToString());
                        if (emp.Contains("Tube_ID"))
                            item.Tube_ID = (emp["Tube_ID"].ToString());
                        if (emp.Contains("FileName"))
                            item.FileName = (emp["FileName"].AsString);
                        if (emp.Contains("Comments"))
                            item.Comments = (emp["Comments"].AsString);
                        if (emp.Contains("OrderName"))
                            item.OrderName = (emp["OrderName"].AsString);
                        if (emp.Contains("Input_Date"))
                            item.Input_Date = (emp["Input_Date"].AsString);

                        #endregion

                        // ClaimReport_Server.Add(item);
                    }
                }
                if (Rack_ID != "")
                {
                    query = new QueryDocument("Rack_ID", Rack_ID);
                    foreach (BsonDocument emp in employees.Find(query))
                    {
                        xmlDataSources item = new xmlDataSources();

                        #region 数据
                        if (emp.Contains("_id"))
                            item.Id = (emp["_id"].ToString());
                        if (emp.Contains("Rack_ID"))
                            item.Rack_ID = (emp["Rack_ID"].ToString());
                        if (emp.Contains("Hole_ID"))
                            item.Hole_ID = (emp["Hole_ID"].ToString());
                        if (emp.Contains("Tube_ID"))
                            item.Tube_ID = (emp["Tube_ID"].ToString());
                        if (emp.Contains("FileName"))
                            item.FileName = (emp["FileName"].AsString);
                        if (emp.Contains("Comments"))
                            item.Comments = (emp["Comments"].AsString);
                        if (emp.Contains("OrderName"))
                            item.OrderName = (emp["OrderName"].AsString);
                        if (emp.Contains("Input_Date"))
                            item.Input_Date = (emp["Input_Date"].AsString);

                        #endregion
                        //  ClaimReport_Server.Add(item);
                    }
                }
                var query1 = Query.And(Query.GTE("Input_Date", start_time.Replace("/", "")), Query.LTE("Input_Date", end_time.Replace("/", "")));
                if (kettext != "" && Rack_ID == "")
                {
                    query1 = Query.And(Query.GTE("Input_Date", start_time.Replace("/", "")), Query.LTE("Input_Date", end_time.Replace("/", "")), Query.Matches("Comments", kettext));
                }
                if (Rack_ID != "" && kettext == "")
                {
                    query1 = Query.And(Query.GTE("Input_Date", start_time.Replace("/", "")), Query.LTE("Input_Date", end_time.Replace("/", "")), Query.EQ("Rack_ID", Rack_ID));
                }
                if (Rack_ID != "" && kettext != "")
                {
                    query1 = Query.And(Query.GTE("Input_Date", start_time.Replace("/", "")), Query.LTE("Input_Date", end_time.Replace("/", "")), Query.EQ("Rack_ID", Rack_ID), Query.EQ("Comments", kettext));
                }



                foreach (BsonDocument emp in employees.Find(query1))
                {
                    xmlDataSources item = new xmlDataSources();
                    #region 数据
                    if (emp.Contains("_id"))
                        item.Id = (emp["_id"].ToString());
                    if (emp.Contains("Rack_ID"))
                        item.Rack_ID = (emp["Rack_ID"].ToString());
                    if (emp.Contains("Hole_ID"))
                        item.Hole_ID = (emp["Hole_ID"].ToString());
                    if (emp.Contains("Tube_ID"))
                        item.Tube_ID = (emp["Tube_ID"].ToString());
                    if (emp.Contains("FileName"))
                        item.FileName = (emp["FileName"].AsString);
                    if (emp.Contains("Comments"))
                        item.Comments = (emp["Comments"].AsString);
                    if (emp.Contains("OrderName"))
                        item.OrderName = (emp["OrderName"].AsString);
                    if (emp.Contains("Input_Date"))
                        item.Input_Date = (emp["Input_Date"].AsString);

                    #endregion

                    ClaimReport_Server.Add(item);
                }


                return ClaimReport_Server;

            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex);
                return null;

                throw ex;
            }
            #endregion
        }
        public void update_OrderServer(List<xmlDataSources> AddMAPResult)
        {

            MongoServer server = MongoServer.Create(connectionString);
            MongoDatabase db1 = server.GetDatabase(DB_NAME);
            MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlData");
            MongoCollection<BsonDocument> employees = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlData");

            //  collection1.RemoveAll();
            if (AddMAPResult == null)
            {
                MessageBox.Show("No Data  input Sever", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            foreach (xmlDataSources item in AddMAPResult)
            {
                IMongoQuery query = Query.EQ("_id", new ObjectId(item.Id));


                #region 集合
                var update = Update.Set("Rack_ID", item.Rack_ID.Trim());
                collection1.Update(query, update);

                if (item.Hole_ID == null)
                    item.Hole_ID = "";

                update = Update.Set("Hole_ID", item.Hole_ID);
                collection1.Update(query, update);

                if (item.Tube_ID == null)
                    item.Tube_ID = "";

                update = Update.Set("Tube_ID", item.Tube_ID);
                collection1.Update(query, update);

                if (item.FileName == null)
                    item.FileName = "";

                update = Update.Set("FileName", item.FileName);
                collection1.Update(query, update);

                if (item.Input_Date == null)
                    item.Input_Date = "";

                update = Update.Set("Input_Date", item.Input_Date);
                collection1.Update(query, update);

                if (item.Comments == null)
                    item.Comments = "";

                update = Update.Set("Comments", item.Comments);
                collection1.Update(query, update);


                if (item.OrderName == null)
                    item.OrderName = "";

                update = Update.Set("OrderName", item.OrderName);
                collection1.Update(query, update);

                #endregion

            }
        }
        public void delete_XMLServer(List<xmlDataSources> AddMAPResult)
        {

            MongoServer server = MongoServer.Create(connectionString);
            MongoDatabase db1 = server.GetDatabase(DB_NAME);
            MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlData");
            MongoCollection<BsonDocument> employees = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlData");
            foreach (xmlDataSources item in AddMAPResult)
            {
                if (item.Id != null && item.Id.Length > 0)
                {
                    IMongoQuery query = Query.EQ("_id", new ObjectId(item.Id));
                    collection1.Remove(query);
                }
            }
        }
        public void deleteall_XMLServer()
        {

            MongoServer server = MongoServer.Create(connectionString);
            MongoDatabase db1 = server.GetDatabase(DB_NAME);
            MongoCollection collection1 = db1.GetCollection("KM_BiotechnologyXML_XmlData");
            MongoCollection<BsonDocument> employees = db1.GetCollection<BsonDocument>("KM_BiotechnologyXML_XmlData");
            collection1.RemoveAll();
        }


    }
}

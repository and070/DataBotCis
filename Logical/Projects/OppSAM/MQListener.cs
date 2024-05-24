using System;
using IBM.WMQ;
using System.Collections;
using System.Threading.Tasks;
using System.IO;
using System.Collections.Generic;
using DataBotV5.Data.Root;

namespace DataBotV5.Logical.Projects.OppSAM
{
    /// <summary>
    /// Clase Logical encargada de MQListener.
    /// </summary>
    class MQListener
    {
        Rooting root = new Rooting();

        public string Put_message(string strInputMsg)
        {
            MQQueueManager queueManager;
            MQQueue queue;
            MQMessage queueMessage;
            MQPutMessageOptions queuePutMessageOptions;

            string QueueManagerName = "QMINTEGRATION";
            string channelName = "SYSTEM.SVRCONN.MIS";
            string connectionName = "10.7.11.26(1414)"; //DEV
            string QueueName = "SAM.OPORTUNIDADES.RQ";
            string message;

            queueManager = new MQQueueManager(QueueManagerName, channelName, connectionName);
            string strReturn = "";
            try
            {
                queue = queueManager.AccessQueue(QueueName, MQC.MQOO_OUTPUT + MQC.MQOO_FAIL_IF_QUIESCING);
                message = strInputMsg;
                queueMessage = new MQMessage();
                queueMessage.WriteString(message);
                queueMessage.Format = MQC.MQFMT_STRING;
                queuePutMessageOptions = new MQPutMessageOptions();
                queue.Put(queueMessage, queuePutMessageOptions);
                strReturn = "OK";//"Message sent to the queue successfully";
            }
            catch (MQException MQexp)
            {
                strReturn = "Exception: " + MQexp.Message;
            }
            catch (Exception exp)
            {
                strReturn = "Exception: " + exp.Message;
            }
            return strReturn;
        }

        public void listener()
        {
            if (root.Mq_mensaje[0] == "0")
            {
                root.Mq_mensaje[0] = "1";
                Task MQ = new Task(conector);
                //if (MQ.Status.ToString() != "Running")
                //{
                MQ.Start();
                //}
            }
        }

        public Action conector = () => 
        {
            Rooting root = new Rooting();

            try
            {
                // root.Mq_mensaje[1] = "";
                List<MQMessage> mqMessage = Listen("QMINTEGRATION", "SAM.OPORTUNIDADES.RQ");
                root.Mq_mensaje[0] = "0";
                //root.Mq_mensaje[1] = mqMessage.ReadString(mqMessage.MessageLength);
                //Save_message(root.Mq_mensaje[1]);
                foreach (MQMessage item in mqMessage)
                {
                Save_message(item.ReadString(item.MessageLength));
                }
                root.Mq_mensaje[2] = "0";
            }
            catch (Exception mqe)
            {
                root.Mq_mensaje[1] = "ERROR: " + mqe.Message.ToString() + " (" + mqe.ToString() + ")";
                Save_message("ERROR: " + mqe.Message.ToString() + " (" + mqe.ToString() + ")");
            }

            List<MQMessage> Listen(string qmName, string queueName)
            {
                List<MQMessage> messages = new List<MQMessage>();
                int openOptions = MQC.MQOO_INPUT_AS_Q_DEF | MQC.MQOO_FAIL_IF_QUIESCING | MQC.MQOO_INQUIRE;

                Hashtable props = new Hashtable();
                props.Add(MQC.HOST_NAME_PROPERTY, "10.7.11.126");        //PRD
                props.Add(MQC.CHANNEL_PROPERTY, "SYSTEM.SVRCONN.MIS");
                props.Add(MQC.PORT_PROPERTY, 1414);
                props.Add(MQC.USER_ID_PROPERTY, "mqm");
                props.Add(MQC.PASSWORD_PROPERTY, "");
                props.Add(MQC.TRANSPORT_PROPERTY, MQC.TRANSPORT_MQSERIES_MANAGED);

                MQQueueManager mqManager = new MQQueueManager(qmName, props);
                MQQueue queue = mqManager.AccessQueue(queueName, openOptions);

                //check si hay, si no esperar a uno
                int depth = queue.CurrentDepth;
                if (depth > 0)
                {
                    for (int i = 0; i < depth; i++)
                    {
                        MQMessage message2 = new MQMessage();
                        queue.Get(message2);
                        messages.Add(message2);
                    }

                }
                else
                {

                    MQGetMessageOptions gmo = new MQGetMessageOptions();
                    gmo.Options = MQC.MQGMO_FAIL_IF_QUIESCING | MQC.MQGMO_WAIT;
                    gmo.WaitInterval = MQC.MQWI_UNLIMITED;
                    MQMessage message = new MQMessage();
                    //wait for message
                    queue.Get(message, gmo);
                    queue.Close();

                    //release resource.
                    mqManager = null;
                    queue = null;
                    gmo = null;
                    System.GC.Collect();
                    messages.Add(message);
                }

                return messages;
            }

            void Save_message(string message)
            {
                string dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\Databot\\OPP_SAM\\";
                System.IO.Directory.CreateDirectory(dir);
                System.Random random = new System.Random();
                string file_name = random.Next(10000).ToString();
                FileInfo fi = new FileInfo(dir + file_name + ".json");
                while (fi.Exists)
                {
                    //fi.Delete();
                    file_name = random.Next(10000).ToString();
                    fi = new FileInfo(dir + file_name + ".json");
                }

                using (StreamWriter sw = fi.CreateText())
                {
                    sw.Write(message);
                }

            }
        };

    }
}

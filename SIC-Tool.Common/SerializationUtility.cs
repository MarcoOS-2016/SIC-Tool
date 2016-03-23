using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.Runtime.Serialization;

namespace SIC_Tool.Common
{
    public class SerializationUtility
    {
        public static string Serialize(object obj)
        {
            return Serialize(obj.GetType(), obj);
        }

        public static string Serialize(Type type, object obj)
        {
            string result = null;
            XmlSerializer ser = new XmlSerializer(type);
            MemoryStream ms = new MemoryStream();

            try
            {
                ser.Serialize(ms, obj);
                int length = (int)ms.Length;
                result = System.Text.Encoding.UTF8.GetString(ms.GetBuffer(), 0, length);
            }
            catch (SerializationException ex)
            {
                throw new SerializationException(
                        string.Format("Serializing the type {0} failed: {1}", type, ex.Message));
            }
            finally
            {
                ms.Close();
            }

            return result;
        }

        public static object DeSerialize(Type type, string msg)
        {
            object result = null;
            MemoryStream ms = null;

            try
            {
                XmlSerializer ser = new XmlSerializer(type);
                ms = new MemoryStream((new UTF8Encoding()).GetBytes(msg));

                result = ser.Deserialize(ms);
            }
            catch (SerializationException ex)
            {
                throw new SerializationException(
                    string.Format("Deserializing the type {0} failed: {1}", type, ex.Message));
            }
            finally
            {
                if (ms != null)
                {
                    ms.Close();
                }
            }

            return result;
        }
    }
}

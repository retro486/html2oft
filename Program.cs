using System;
using System.Collections.Generic;
using System.Text;
using OpenMcdf;
using System.IO;

namespace HTML2OFT
{
	class MainClass
	{
        static string TEMPLATE_FILENAME = "Blank.oft";
        static string PROPERTIES_STREAM_ID = "__properties_version1.0";
        
        static string HTML_STREAM_ID = "__substg1.0_1013001E";
        static string PLAINTEXT_STREAM_ID = "__substg1.0_1000001E";
        static string COMPRESSED_RTF_STREAM_ID = "__substg1.0_10090102";
        
        static byte[] HTML_ID = { 0x10, 0x13, 0x00, 0x1e };
        static byte[] PLAINTEXT_ID = { 0x10, 0x00, 0x00, 0x1e };
        static byte[] COMPRESSED_RTF_ID = { 0x10, 0x09, 0x01, 0x02 };
        static byte[] RTF_IN_SYNC_ID = { 0x0e, 0x1f, 0x00, 0x0b };
        static byte[] NATIVE_BODY_ID = { 0x10, 0x16, 0x00, 0x03 };

        static byte[] CLEAR_WORD = { 0x00, 0x00, 0x00, 0x00 };
        static byte[] RW_FLAG = { 0x06, 0x00, 0x00, 0x00 };
        static byte[] ENABLED_FLAG = { 0x03, 0x00, 0x00, 0x00 };

        public static void Main (string[] args)
		{
			if (args.Length < 2) {
				Console.WriteLine ("Must specify input HTML then output OFT filenames.");
				return;
			}

			string inputFilename = args[0];
			string outputFilename = args[1];
            byte[] properties;

            File.Delete(outputFilename);

            CompoundFile cf = new CompoundFile(TEMPLATE_FILENAME);

            StreamReader in_file = File.OpenText(inputFilename);
            string in_data = in_file.ReadToEnd();
            in_file.Close();

            CFStream s = null;

            try
            {
                cf.RootStorage.Delete(PLAINTEXT_STREAM_ID);
            }
            catch (CFItemNotFound e) { }
            //s = cf.RootStorage.AddStream(PLAINTEXT_STREAM_ID);
            //s.SetData(StringToBytes(in_data));

            try
            {
                cf.RootStorage.Delete(HTML_STREAM_ID);
            }
            catch (CFItemNotFound e) { }
            //s = cf.RootStorage.AddStream(HTML_STREAM_ID);
            //s.SetData(StringToBytes(in_data));

            try
            {
                cf.RootStorage.Delete(COMPRESSED_RTF_STREAM_ID);
            }
            catch (CFItemNotFound e) { }
            s = cf.RootStorage.AddStream(COMPRESSED_RTF_STREAM_ID); // leave empty
            byte[] rtfData = RTFTools.CompressString(RTFTools.AttachRTFHeader(in_data));
            s.SetData(rtfData);
                        
            try
            {
                s = cf.RootStorage.GetStream(PROPERTIES_STREAM_ID);
                properties = s.GetData();

                properties = SetPropertyValue(properties, COMPRESSED_RTF_ID, RW_FLAG, BitConverter.GetBytes((Int32)rtfData.Length), ENABLED_FLAG);
                
                // Try to make use of Best Body Algorithm: http://msdn.microsoft.com/en-us/library/hh369831%28v=exchg.80%29.aspx
                properties = SetPropertyValue(properties, RTF_IN_SYNC_ID, RW_FLAG, new byte[]{0x01, 0x00, 0x00, 0x00}, CLEAR_WORD);
                properties = SetPropertyValue(properties, NATIVE_BODY_ID, RW_FLAG, new byte[] { 0x02, 0x00, 0x00, 0x00 }, CLEAR_WORD); // Set native body to RTF Compressed

                // Clear other formats
                properties = SetPropertyValue(properties, HTML_ID, RW_FLAG, CLEAR_WORD, ENABLED_FLAG);
                properties = SetPropertyValue(properties, PLAINTEXT_ID, RW_FLAG, CLEAR_WORD, ENABLED_FLAG);

                s.SetData(properties);
            }
            catch (OpenMcdf.CFItemNotFound e)
            {
                Console.WriteLine("FAILURE: Unable to update properties. OFT will be corrupt.");
            }

            cf.Save(outputFilename);
			cf.Close ();

			Console.WriteLine ("Done generating OFT. Any errors will be reported above.");
            Console.ReadLine();
		}

		protected static byte[] StringToBytes(string input) {
            return Encoding.ASCII.GetBytes(input.ToCharArray());
		}

        protected static byte[] SetPropertyValue(byte[] properties, byte[] propertyId, byte[] flags, byte[] length, byte[] value) {
            // http://blogs.msdn.com/b/openspecification/archive/2009/11/06/msg-file-format-part-1.aspx
            // Properties in stream are every 16 bytes, first 4 bytes are ID in little endian. propertyId must be 4 bytes and value must be 4 bytes.
            // Convert the given property to little endian
            byte[] littleId = new byte[propertyId.Length];
            int j = 3;
            for (int i = 0; i < 4; i+=2)
            {
                littleId[j--] = propertyId[i];
                littleId[j--] = propertyId[i + 1];
            }

            for (int i = 0; i < properties.Length; i += 16)
            {
                byte[] singleId = new byte[4];
                Buffer.BlockCopy(properties, i, singleId, 0, 4);
                byte[] singleFlag = new byte[4];
                Buffer.BlockCopy(properties, i + 4, singleFlag, 0, 4);
                byte[] singleLength = new byte[4];
                Buffer.BlockCopy(properties, i + 8, singleLength, 0, 4);
                byte[] singleVal = new byte[4];
                Buffer.BlockCopy(properties, i + 12, singleVal, 0, 4);

                if (UglyByteArraysAreEqual(singleId, littleId))
                {
                    Buffer.BlockCopy(flags, 0, properties, i + 4, 4);
                    Buffer.BlockCopy(length, 0, properties, i + 8, 4);
                    Buffer.BlockCopy(value, 0, properties, i + 12, 4);
                    return properties;
                }
            }

            Console.WriteLine("Unable to find propertyId: " + BitConverter.ToString(propertyId) + ", so it was added.");
            int originalLength = properties.Length;
            byte[] growProperties = new byte[originalLength + 16];
            Buffer.BlockCopy(properties, 0, growProperties, 0, originalLength);
            Buffer.BlockCopy(littleId, 0, growProperties, originalLength, 4);
            Buffer.BlockCopy(flags, 0, growProperties, originalLength + 4, 4);
            Buffer.BlockCopy(length, 0, growProperties, originalLength + 8, 4);
            Buffer.BlockCopy(value, 0, growProperties, originalLength + 12, 4);
            return growProperties;
        }

        protected static bool UglyByteArraysAreEqual(byte[] a, byte[] b) {
            if (a.Length != b.Length) return false;

            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) return false;
            
            return true;
        }
	}
}

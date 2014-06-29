using System;
using System.Collections.Generic;
using System.Text;
using OpenMcdf;
using System.IO;

namespace HTML2OFT
{
	class MainClass
	{
		static string HTML_STREAM_ID = "__substg1.0_1013001E";
        static string HTML_UNICODE_STREAM_ID = "__substg1.0_1013001F";
        static string PLAINTEXT_STREAM_ID = "__substg1.0_1000001F";
        static string PLAINTEXT_UNICODE_STREAM_ID = "__substg1.0_1000001E";
        static string COMPRESSED_RTF_STREAM_ID = "__substg1.0_10090102";
        static string SUBJECT_STREAM_ID = "__substg1.0_0037001E";
        static string PROPERTIES_STREAM_ID = "__properties_version1.0";
        static string TEMPLATE_FILENAME = "Blank.oft";

        static byte[] CLEAR_WORD = { 0x00, 0x00, 0x00, 0x00 };
        static byte[] COMPRESSED_RTF_ID = { 0x10, 0x09, 0x01, 0x02 };
        static byte[] PLAINTEXT_ID = { 0x10, 0x00, 0x00, 0x1e };
        static byte[] HTML_ID = { 0x10, 0x13, 0x00, 0x1e };
        static byte[] STOCK_TYPE = { 0x06, 0x00, 0x00, 0x00 };
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
            
			try {
				s = cf.RootStorage.GetStream(HTML_STREAM_ID);
			} catch(OpenMcdf.CFItemNotFound e) {
				Console.WriteLine ("Warning: HTML stream not found. Creating it...");
				s = cf.RootStorage.AddStream (HTML_STREAM_ID);
			}
            s.SetData(StringToBytes(in_data));

            try
            {
                cf.RootStorage.Delete(PLAINTEXT_UNICODE_STREAM_ID);
            }
            catch (CFItemNotFound e) { }
            try
            {
                cf.RootStorage.Delete(HTML_UNICODE_STREAM_ID);
            }
            catch (CFItemNotFound e) { }
            try
            {
                cf.RootStorage.Delete(COMPRESSED_RTF_STREAM_ID);
            }
            catch (CFItemNotFound e) { }

            //try
            //{
            //    s = cf.RootStorage.GetStream(PLAINTEXT_STREAM_ID);
            //}
            //catch (OpenMcdf.CFItemNotFound e)
            //{
            //    Console.WriteLine("Warning: PLAINTEXT stream not found. Creating it...");
            //    s = cf.RootStorage.AddStream(PLAINTEXT_STREAM_ID);
            //}
            //s.SetData(StringToBytes("Hello there!"));

            //try
            //{
            //    s = cf.RootStorage.GetStream(PLAINTEXT_UNICODE_STREAM_ID);
            //}
            //catch (OpenMcdf.CFItemNotFound e)
            //{
            //    Console.WriteLine("Warning: PLAINTEXT Unicode stream not found. Creating it...");
            //    s = cf.RootStorage.AddStream(PLAINTEXT_UNICODE_STREAM_ID);
            //}
            //s.SetData(StringToBytes("Hello there!"));

            //try
            //{
            //    s = cf.RootStorage.GetStream(HTML_UNICODE_STREAM_ID);
            //}
            //catch (OpenMcdf.CFItemNotFound e)
            //{
            //    Console.WriteLine("Warning: HTML stream not found. Creating it...");
            //    s = cf.RootStorage.AddStream(HTML_UNICODE_STREAM_ID);
            //}
            //s.SetData(StringToBytes(in_data));
            
            try
            {
                s = cf.RootStorage.GetStream(PROPERTIES_STREAM_ID);
                properties = s.GetData();

                // Clear RTF property
                properties = SetPropertyValue(properties, COMPRESSED_RTF_ID, STOCK_TYPE, CLEAR_WORD, ENABLED_FLAG);
                // Clear PLAINTEXT property
                properties = SetPropertyValue(properties, PLAINTEXT_ID, STOCK_TYPE, CLEAR_WORD, ENABLED_FLAG);
                // Set HTML property
                properties = SetPropertyValue(properties, HTML_ID, STOCK_TYPE, BitConverter.GetBytes((Int32)in_data.Length), ENABLED_FLAG);

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

        // Eg to clear RTF content (0x10090102): [0x10, 0x09, 0x01, 0x02], [0x06, 0x00, 0x00, 0x00], [0x00, 0x00, 0x00, 0x00], [0x00, 0x00, 0x00, 0x00]
        // And set HTML content (0x1013001E): [0x10, 0x13, 0x00, 0x1e], [0x06, 0x00, 0x00, 0x00], [0x50, 0x00, 0x00, 0x00], [0x03, 0x00, 0x00, 0x00] (if the length is 80 bytes [0x50 in hex])
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
            Console.WriteLine("Looked for: " + BitConverter.ToString(littleId));
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

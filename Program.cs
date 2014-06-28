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
		static string TEMPLATE_FILENAME = "Blank.oft";


		public static void Main (string[] args)
		{
			if (args.Length < 2) {
				Console.WriteLine ("Must specify input HTML then output OFT filenames.");
				return;
			}

			string INPUT_FILENAME = args[0];
			string OUTPUT_FILENAME = args[1];

			File.Delete (OUTPUT_FILENAME);

			CompoundFile cf = new CompoundFile(TEMPLATE_FILENAME);
			StreamReader in_file = File.OpenText (INPUT_FILENAME);

			CFStream s = null;
			try {
				s = cf.RootStorage.GetStream(HTML_STREAM_ID);
			} catch(OpenMcdf.CFItemNotFound e) {
				Console.WriteLine ("Warning: HTML stream not found. Creating it...");
				s = cf.RootStorage.AddStream (HTML_STREAM_ID);
			}

			s.AppendData(StringToBytes(in_file.ReadToEnd()));

			in_file.Close ();
			cf.Save (OUTPUT_FILENAME);
			cf.Close ();

			Console.WriteLine ("Done generating OFT. Any errors will be reported above.");
		}

		protected static byte[] StringToBytes(string input) {
			char[] chars = input.ToCharArray ();
			List<byte> bytes = new List<byte> ();

			for (int i = 0; i < chars.Length; i++) {
				bytes.AddRange (BitConverter.GetBytes (chars [i]));
			}

			return bytes.ToArray ();
		}
	}
}

// rtf_test.cpp : Defines the entry point for the console application.
//

//https://stackoverflow.com/questions/9036568/best-way-to-remove-extra-marks-from-an-rtf-text
//https://stackoverflow.com/questions/1046824/how-can-you-append-a-line-not-a-par-to-a-system-windows-forms-richtextbox

//formating information
//http://www.pindari.com/

//the doc............
//https://www.microsoft.com/en-us/download/details.aspx?id=10725

#include "stdafx.h"
#include "richtext_io.h"

const char _RichEditPreamble[]= R"({\rtf1\ansi\deff0\deflang1033{\fonttbl
{\f0\froman\fprq2 Times New Roman;}{\f1\froman\fprq2\fcharset0 Times New Roman;}
{\f2\fmodern\fprq1 Courier New;}{\f1\fnil\fcharset2 Symbol;}}
{\colortbl\red0\green0\blue0;\red129\green129\blue129;
\red128\green0\blue0;\red0\green128\blue0;}
\viewkind4\uc1\pard\f0\fs24\sa120 )";

extern const char* cell_test;
void rtf_to_file( COleRichEditCtrl& rtf, char* filename)
{
	auto buf = GetRichText(rtf);
	std::ofstream st(filename);
	st << buf.get();
	st.close();
}

int main()
{
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))  
	{       
		// TODO: change error code to suit your needs       
		std::cout << _T("Fatal Error: MFC initialization failed") << std::endl;       
		return 1;   
	}   
	CoInitialize(NULL);

	//test 1
	{
		COleRichEditCtrl rtf;
		rtf.Create();
		//the next two lines will make no difference in the output.
		rtf.SetSel(0,-1);
		rtf.ReplaceSel("");

		auto result = GetRichText(rtf);
		std::cout << "result of empty control:\n"
		<< result.get() << "****\n\n";
		//open in wordpad and there is no paragraph/end of line marker to edit
		//even though the file ends with \par. The same goes for Open Office and MSWord
		std::ofstream ts("..\\rtf_1.rtf");
		ts << result.get();
	}

	//test 2
	{
		COleRichEditCtrl rtf;
		rtf.Create();
		std::string source(_RichEditPreamble);
		//the following gets the same result wheather there is 'no' or 'one \par' at end.
		//the output will have just one '\par' which is considered a line that has no paragraph marker
		//with two '\par' control words at end you will see a new line in the editors.
		//Wordpad, Open Office, and MSWord
		source += R"(\cf1 test 1\par\par})";
		SetRichText(rtf,source.c_str());
		auto result = GetRichText(rtf);
		std::cout << "result of added to control:\n"
			<< result.get() << "****\n\n";
		std::ofstream ts("..\\rtf_2.rtf");
		ts << result.get();
	}

	//test 3 concatenate
	{
		COleRichEditCtrl rtf;
		rtf.Create();
		std::string source(_RichEditPreamble);
		//there must be two \par to retain a \par for the concatenate
		source += R"(\cf2\b\fs32\pncf0\ulwave test 1\par\par})";
		SetRichText(rtf,source.c_str());

		std::string source2(_RichEditPreamble);
		//only one par here to get an ending par in the doc.
		source2+= R"(\cf0 test 2\par})";

		rtf.SetSel(-1,-1);
		SetRichText(rtf,source2.c_str(), SF_RTF | SFF_SELECTION);

		auto result = GetRichText(rtf);
		//but, the result is that char format from the first insertion
		//follows at the end of the doc now. If the client adds text
		//it won't be with the expected font.
		std::cout << "result of concatenated to control:\n"
			<< result.get() << "****\n\n";
		std::ofstream ts("..\\rtf_3.rtf");
		ts << result.get();

		//so........
		auto len = rtf.GetTextLength();
		std::cout << "len: " << len << "\n\n";
		rtf.SetSel(len - 3, -1);  //why 3?
		//rtf.ReplaceSel("\r\n");
		//or
		SetRichText(rtf,R"({\rtf1\par})", SF_RTF | SFF_SELECTION);
		auto result2 = GetRichText(rtf);

		std::cout << "result of modified concatenated to control:\n"
			<< result2.get() << "****\n\n";
		std::ofstream ts2("..\\rtf_4.rtf");
		ts2 << result2.get();
	}
	//TEST call to AppendRichText
	{
		COleRichEditCtrl rtf;
		rtf.Create();
		std::string source(_RichEditPreamble);
		source += R"(\cf2\uline First Line\par\par)";
		SetRichText(rtf,source.c_str());
		//arrrrrrrrrr.......
		rtf_to_file(rtf,"..\\xxx.rtf");
		//{
		//	auto buf = GetRichText(rtf);
		//	std::ofstream st("..\\xxx.rtf");
		//	st << buf.get();
		//	st.close();
		//}
		source = _RichEditPreamble;
		source+= R"(\cf0\uline0 This test\par)";
		AppendRichText(rtf,source.c_str());
		rtf_to_file(rtf,"..\\xxxxx.rtf");

		auto result = GetRichText(rtf);
		std::cout << "result of modified concatenated to control:\n"
			<< result.get() << "****\n\n";
		std::ofstream ts("..\\rtf_5.rtf");
		ts << result.get();
	}
	//cell test
	{
		std::string source(_RichEditPreamble);
		source += cell_test;
		std::ofstream ts2("..\\rtf_cells.rtf");
		ts2 << source;

	}
	return 0;
}

const char* cell_test = R"({\pard Hmmm \par}
{\pard
\trowd\trgaph300\trleft400\cellx1500\cellx3000\cellx4500
\pard\intbl Wun. Doo wah ditty ditty dum ditty do \cell
\pard\intbl Too. Doo wah ditty ditty dum ditty do \cell
\pard\intbl Chree. Doo wah ditty ditty dum ditty do \cell
\row
\trowd\trgaph300\trleft400\cellx1500\cellx3000\cellx4500
\pard\intbl Foh. Doo wah ditty ditty dum ditty do \cell
\pard\intbl Fahv. Doo wah ditty ditty dum ditty do \cell
\pard\intbl Saxe. Doo wah ditty ditty dum ditty do \cell
\row
\trowd\trgaph300\trleft400\cellx1500\cellx3500
\pard\intbl Saven. Doo wah ditty ditty dum ditty do \cell
\pard\intbl Ight. Doo wah ditty ditty dum ditty do \cell
\row
}
{\pard I LIKE PIE})";

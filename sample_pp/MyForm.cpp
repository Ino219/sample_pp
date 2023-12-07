#include "MyForm.h"

using namespace samplepp;
using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;
using namespace Microsoft::Office::Interop::Outlook;

using namespace System::Collections::Generic;


[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void samplepp::MyForm::MyForm_DragDrop(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//outlookのインスタンス作成
	Microsoft::Office::Interop::Outlook::Application^ ol = gcnew Microsoft::Office::Interop::Outlook::ApplicationClass();
	MailItem^ mailItem = (MailItem^)ol->CreateItem(OlItemType::olMailItem);
	if (mailItem != nullptr)
	{
		// To
		Recipient^ to = mailItem->Recipients->Add("XXX@XXX.co.jp");
		to->Type = (int)Microsoft::Office::Interop::Outlook::OlMailRecipientType::olTo;

		// Cc
		Recipient^ cc = mailItem->Recipients->Add("YYY@YYY.co.jp");
		cc->Type = (int)Microsoft::Office::Interop::Outlook::OlMailRecipientType::olCC;

		// アドレス帳の表示名で表示できる
		mailItem->Recipients->ResolveAll();

		// 件名
		mailItem->Subject = "件名";

		//添付ファイル
		mailItem->Attachments->Add("C:\\Users\\chach\\Desktop\\test.pptx",OlAttachmentType::olByValue,1, "C:\\Users\\chach\\Desktop\\test.pptx");

		// 本文
		mailItem->Body = "本文";

		// 表示(Displayメソッド引数のtrue/falseでモーダル/モードレスウィンドウを指定して表示できる)
		//mailItem->Display(true);

		mailItem->SaveAs("C:\\Users\\chach\\Desktop\\test.msg",OlSaveAsType::olMSG);
	}
	else {
		MessageBox::Show("t");
			
	}

	//フォームへのドラッグドロップイベント
	array<String^>^ file = (array<String^>^)e->Data->GetData(DataFormats::FileDrop, false);
	//拡張子の取得
	String^	extension = System::IO::Path::GetExtension(file[0]);
	//フルパス取得
	String^ path = System::IO::Path::GetFullPath(file[0]);

	Microsoft::Office::Interop::PowerPoint::Application^ app_ = nullptr;
	List<Microsoft::Office::Interop::PowerPoint::Shape^>^ shapeList = gcnew List<Microsoft::Office::Interop::PowerPoint::Shape^>;

	if (extension == ".pptx") {

		String^ files=System::IO::Path::GetFileName(file[0]);
		if (files == "基板ファイル名_レイアウト設計報告書.pptx") {
			MessageBox::Show("OK");
		}
		
		//パワーポイントファイルの場合、処理を開始
		app_ = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
		Microsoft::Office::Interop::PowerPoint::Presentations^ presen = app_->Presentations;
		Microsoft::Office::Interop::PowerPoint::Presentation^ presense = presen->Open(
			path,
			MsoTriState::msoFalse,
			MsoTriState::msoFalse,
			MsoTriState::msoFalse
		);

		//スライドを画像として出力
		int width_ = (int)presense->PageSetup->SlideWidth;
		int height_ = (int)presense->PageSetup->SlideHeight;
		String^ file2;

		for (int i = 1; i <= presense->Slides->Count; i++) {
			// JPEGとして保存
			file2 = "C:\\Users\\chach\\Desktop\\ffolder\\" + String::Format("\slide{0:0000}.png", i);
			presense->Slides[i]->Export(file2, "png", width_, height_);
		}
		//保存した画像を取得して、パワーポイントに添付
		presense->Slides->Add(presense->Slides->Count,Microsoft::Office::Interop::PowerPoint::PpSlideLayout::ppLayoutBlank);
		//LinkToFileとSaveWidthDocumentはどちらかをTrueにする
		//今回はスライドを画像出力した後、消去するので、LinkToFileをTrueにすると、エラーになり、表示ができないため、後者をtrueにする
		presense->Slides[presense->Slides->Count]->Shapes->AddPicture(file2, Microsoft::Office::Core::MsoTriState::msoFalse, Microsoft::Office::Core::MsoTriState::msoTrue, 10.0, height_/3, width_/2.3, height_/2.3);
		
		//固定値の列数
		int c = 3;
		//各列に代入する配列
		cli::array<String^>^ testone = gcnew cli::array<String^>(4){"a", "b", "c", "d" };
		cli::array<String^>^ testtwo = gcnew cli::array<String^>(4) { "f", "g", "h", "i" };
		cli::array<String^>^ testthree = gcnew cli::array<String^>(4) { "k", "l", "m", "n" };
		//行数を定義
		int r = testone->Length+1;
		//テーブルを追加
		Microsoft::Office::Interop::PowerPoint::Shape^ tab = presense->Slides[presense->Slides->Count]->Shapes->AddTable(r,c, width_ / 2, height_ / 3, width_ / 3, height_ / 3);
		for (int i = 0; i < r*c; i++) {
			//ヘッダー	
			switch (i) {
				case 0:
					tab->Table->Columns[1]->Cells[1]->Shape->TextFrame->TextRange->Text = "aa";
					break;
				case 1:
					tab->Table->Columns[2]->Cells[1]->Shape->TextFrame->TextRange->Text = "bb";
					break;
				case 2:
					tab->Table->Columns[3]->Cells[1]->Shape->TextFrame->TextRange->Text = "cc";
					break;
			}
			//ヘッダー以降の代入
			if (i > 2) {
				//1列目のセルであれば、testone配列の値を取得
				if ((i + 1) % 3 == 1) {
					tab->Table->Columns[(i + 1) % 3]->Cells[i/c+1]->Shape->TextFrame->TextRange->Text = testone[i/c-1];
				}
				//2列目のセルであれば、testtwo配列の値を取得
				else if ((i + 1) % 3 == 2) {
					tab->Table->Columns[(i + 1) % 3]->Cells[i / c + 1]->Shape->TextFrame->TextRange->Text = testtwo[i/c-1];
				}
				//3列目のセルであれば、testthree配列の値を取得
				else if ((i + 1) % 3 == 0) {
					tab->Table->Columns[c]->Cells[i / c + 1]->Shape->TextFrame->TextRange->Text = testthree[i/c-1];
				}
			}
		}
		
		//ファイルを保存
		presense->SaveCopyAs("C:\\Users\\chach\\Desktop\\ffolder\\t.pptx", Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsPresentation, Microsoft::Office::Core::MsoTriState::msoFalse);
		//スライド画像を消去
		//System::IO::File::Delete(file2);

		//図形名
		String^ name;
		//テキストフレーム
		Microsoft::Office::Interop::PowerPoint::TextFrame^ text;
		//テキスト内容
		String^ text2;
		//画像パス
		String^ picPath;
		//フォント名
		String^ font;
		//フォントサイズ
		int TextSize;
		//パネル内項目数
		int itemCount = 0;
		//テキストかどうか
		bool textCheck;
		//画像かどうか
		bool pictureCheck;
		//表かどうか
		bool tableCheck;
		//図形スタイル
		MsoShapeStyleIndex style;
		//図形タイプ
		MsoShapeType type;
		int height;
		int width;
		int x;
		int y;

		
		//picturebox表示用
		Bitmap^ b;

		//タイトルがあれば、以下のやり方でタイトル情報を取得できる
		if (presense->Slides[1]->Shapes->HasTitle == MsoTriState::msoTrue) {
			String^ titleText = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Text;
			String^ titlefont = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Font->Name;
			String^ titlefont2 = presense->Slides[1]->Shapes->Title->AlternativeText;
			int title_size = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Font->Size;
		}
		

		

		//presense->Slides[1]->Shapes->Title->Export("C:\\Users\\chach\\Desktop\\sample_title.jpg", PpShapeFormat::ppShapeFormatJPG, width, height, PpExportMode::ppClipRelativeToSlide);

		//図形を一つずつ取得し、処理
		for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in presense->Slides[1]->Shapes)
		{
			//スライド表示用のキャンバス幅
			int slideWidth = presense->PageSetup->SlideWidth;
			//スライド表示用のキャンバス高さ
			int slideHeight = presense->PageSetup->SlideHeight;
			pictureBox1->Width = slideWidth;
			pictureBox1->Height = slideHeight;
			//描画用ビットマップ
			b = gcnew Bitmap(pictureBox1->Width, pictureBox1->Height);
			//真偽値変数初期化
			pictureCheck = false;
			textCheck = false;
			//図形の名前を取得
			name = var->Name;
			//図形のタイプを取得
			style = var->ShapeStyle;
			type = var->Type;
			//図形の高さ
			height = var->Height;
			//図形の幅
			width = var->Width;
			//図形の位置
			x = var->Left;
			y = var->Top;

			

			//テキストタイプの処理
			if (var->HasTextFrame==MsoTriState::msoTrue) {
				textCheck = true;
				//テキストが含まれている場合
				text = var->TextFrame;
				//テキスト取得
				text2 = text->TextRange->Text;
				//テキストサイズ
				TextSize = text->TextRange->Font->Size;
				//テキストフォント名
				font = var->AlternativeText;
				//var->Export("C:\\Users\\chach\\Desktop\\sample_text"+text+".jpg", PpShapeFormat::ppShapeFormatJPG, width, height, PpExportMode::ppScaleXY);
			}

			if (var->HasTable == MsoTriState::msoTrue) {
				pictureCheck = true;
				//テーブルを保有していた場合、一旦は画像として保持
				var->Export("C:\\Users\\chach\\Desktop\\sample_hyou.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleXY);
				picPath = "C:\\Users\\chach\\Desktop\\sample_hyou.bmp";
			}
			//埋め込みデータの場合
			if (type==MsoShapeType::msoEmbeddedOLEObject) {
				pictureCheck = true;
				var->Export("C:\\Users\\chach\\Desktop\\sample_umekomi.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleToFit);
				picPath = "C:\\Users\\chach\\Desktop\\sample_umekomi.bmp";
			}
			//画像ファイルの場合
			if (type == MsoShapeType::msoPicture) {
				pictureCheck = true;
				var->Export("C:\\Users\\chach\\Desktop\\sample_pic.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleXY);
				picPath = "C:\\Users\\chach\\Desktop\\sample_pic.bmp";
			}
			//図形構造体に代入
			Ashapes^ temp = gcnew Ashapes;
			temp->name = name;
			temp->t = type;
			temp->height = height;
			temp->width = width;
			temp->x = x;
			temp->y = y;
			temp->text = textCheck;
			temp->textVal = text2;
			temp->picture = pictureCheck;
			temp->picturePath = picPath;
				
			//
			//図形情報を取得
			shapeList->Add(var);
			shapesList->Add(temp);
		}
		int count = 0;
		//取得した図形情報をパネルとpictureBoxに表示処理していく
		for each (Ashapes^ var in shapesList)
		{
			//画像の場合
			if (var->picture) {
				//パネル表示用のピクチャーボックス
				pic = gcnew PictureBox;
				//コントロールに追加
				this->Controls->Add(pic);
				//ピクチャーボックスの座標
				pic->Top = 25 * itemCount;
				//高さ
				pic->Height = 25;
				//画像パスからイメージ作成
				System::Drawing::Image^ img= System::Drawing::Image::FromFile(var->picturePath);
				//イメージのはめ込み
				pic->Image = img;
				//表示用のピクチャーボックスのイメージを空のビットマップに設定
				pictureBox1->Image = b;
				//グラフィックを設定
				Graphics^ gr = Graphics::FromImage(pictureBox1->Image);
				//画像表示
				gr->DrawImage(img, var->x, var->y, var->width, var->height);
				//描画更新
				pictureBox1->Refresh();
				count++;
				pic->MouseDown += gcnew System::Windows::Forms::MouseEventHandler(this, &MyForm::picture_MouseDown);
				
			}
			else {
				//ラベル作成
				tx = gcnew Label;
				//コントロールに追加
				this->Controls->Add(tx);
				tx->Top = 25 * itemCount;
				tx->Text = var->textVal;
				//ピクチャーボックス作成
				//空のビットマップに文字描画
				//フォントなど指定
				tx->MouseDown += gcnew System::Windows::Forms::MouseEventHandler(this, &MyForm::text_MouseDown);
			}
			itemCount++;
		}
		//パワーポイントのインスタンス処理
		presense->Close();
		app_->Quit();

	}
	return System::Void();
}

System::Void samplepp::MyForm::MyForm_DragEnter(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//フォームへのファイルドラッグドロップイベント
	if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		e->Effect = DragDropEffects::All;
	}
	else {
		e->Effect = DragDropEffects::None;
	}
	return System::Void();
}

System::Void samplepp::MyForm::picture_DragDrop(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//パネルからのドラッグドロップイベント
	if (e->Data->GetDataPresent(DataFormats::Bitmap)) {
		System::Drawing::Graphics^ tegr = pictureBox1->CreateGraphics();
		Image^ img = (Image^)e->Data->GetData(DataFormats::Bitmap);
		int xp = pictureBox1->MousePosition.X;
		int yp = pictureBox1->MousePosition.Y;
		tegr->DrawImage((Image^)e->Data->GetData(DataFormats::Bitmap), xp, yp);
		//pictureBox1->Image = (Image^)e->Data->GetData(DataFormats::Bitmap);

	}
	else if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		//データオブジェクトの受け取り
		array<String^>^ temp=(array<String^>^)e->Data->GetData(DataFormats::FileDrop);
		
	}
	return System::Void();
}

System::Void samplepp::MyForm::picture_DragEnter(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		e->Effect = DragDropEffects::All;
	}
	else {
		e->Effect = DragDropEffects::None;
	}
	return System::Void();
}

System::Void samplepp::MyForm::picture_MouseDown(System::Object ^ sender, System::Windows::Forms::MouseEventArgs ^ e)
{
	//マウスクリックしたピクチャーボックスを取得
	pic = (PictureBox^)sender;
	//ドラッグドロップイベントを設定
	pic->DoDragDrop(pic->Image, DragDropEffects::All);
	return System::Void();
}

System::Void samplepp::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	//ロード時、ドロップイベントを許可
	pictureBox1->AllowDrop = true;
	//イベント駆動を設定
	pictureBox1->DragDrop += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::picture_DragDrop);
	pictureBox1->DragEnter += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::picture_DragDrop);
	return System::Void();
}


System::Void samplepp::MyForm::text_MouseDown(System::Object ^ sender, System::Windows::Forms::MouseEventArgs ^ e)
{
	//パネルからラベルを選択した場合の処理
	text_pickUp = true;
	//必要なデータの配列を作成
	array<String^>^ data = gcnew array<String^>(3);
	//どのラベルが選択されたか
	tx = (Label^)sender;
	//ラベルのテキストを取得
	String^ lbtext = tx->Text;
	//テキストと一致する図形情報を検索し、テキスト内容、フォント、サイズを取得し、配列に格納
	for each (Ashapes^ var in shapesList)
	{
		if (var->text&&(lbtext == var->textVal)) {
			data[0] = var->textVal;
			data[1] = var->fontName;
			data[2] = var->textSize.ToString();
		}
	}
	//データオブジェクトとしてまとめ、ファイルドロップ形式で紐づける
	DataObject^ dobj = gcnew DataObject(DataFormats::FileDrop, data);
	//ドラッグドロップイベントでデータオブジェクトを引っ張る
	tx->DoDragDrop(dobj, DragDropEffects::All);
	return System::Void();
}


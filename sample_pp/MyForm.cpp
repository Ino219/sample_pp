#include "MyForm.h"

using namespace samplepp;
using namespace Microsoft::Office::Core;
using namespace Microsoft::Office::Interop::PowerPoint;
using namespace System::Collections::Generic;


[STAThreadAttribute]

int main() {
	System::Windows::Forms::Application::Run(gcnew MyForm());
	return 0;
}

System::Void samplepp::MyForm::MyForm_DragDrop(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//�t�H�[���ւ̃h���b�O�h���b�v�C�x���g
	array<String^>^ file = (array<String^>^)e->Data->GetData(DataFormats::FileDrop, false);
	//�g���q�̎擾
	String^	extension = System::IO::Path::GetExtension(file[0]);
	//�t���p�X�擾
	String^ path = System::IO::Path::GetFullPath(file[0]);

	Microsoft::Office::Interop::PowerPoint::Application^ app_ = nullptr;
	List<Microsoft::Office::Interop::PowerPoint::Shape^>^ shapeList = gcnew List<Microsoft::Office::Interop::PowerPoint::Shape^>;

	if (extension == ".pptx") {

		

		
		//�p���[�|�C���g�t�@�C���̏ꍇ�A�������J�n
		app_ = gcnew Microsoft::Office::Interop::PowerPoint::ApplicationClass();
		Microsoft::Office::Interop::PowerPoint::Presentations^ presen = app_->Presentations;
		Microsoft::Office::Interop::PowerPoint::Presentation^ presense = presen->Open(
			path,
			MsoTriState::msoFalse,
			MsoTriState::msoFalse,
			MsoTriState::msoFalse
		);

		//�X���C�h���摜�Ƃ��ďo��
		int width_ = (int)presense->PageSetup->SlideWidth;
		int height_ = (int)presense->PageSetup->SlideHeight;
		String^ file2;

		for (int i = 1; i <= presense->Slides->Count; i++) {
			// JPEG�Ƃ��ĕۑ�
			file2 = "C:\\Users\\chach\\Desktop\\ffolder\\" + String::Format("\slide{0:0000}.jpg", i);
			presense->Slides[i]->Export(file2, "jpg", width_, height_);
		}
		//�ۑ������摜���擾���āA�p���[�|�C���g�ɓY�t
		presense->Slides->Add(presense->Slides->Count,Microsoft::Office::Interop::PowerPoint::PpSlideLayout::ppLayoutBlank);
		//LinkToFile��SaveWidthDocument�͂ǂ��炩��True�ɂ���
		//����̓X���C�h���摜�o�͂�����A��������̂ŁALinkToFile��True�ɂ���ƁA�G���[�ɂȂ�A�\�����ł��Ȃ����߁A��҂�true�ɂ���
		presense->Slides[presense->Slides->Count]->Shapes->AddPicture(file2, Microsoft::Office::Core::MsoTriState::msoFalse, Microsoft::Office::Core::MsoTriState::msoTrue, 10.0, height_/3, width_/2.3, height_/2.3);
		//�t�@�C����ۑ�
		presense->SaveCopyAs("C:\\Users\\chach\\Desktop\\ffolder\\t.pptx", Microsoft::Office::Interop::PowerPoint::PpSaveAsFileType::ppSaveAsPresentation, Microsoft::Office::Core::MsoTriState::msoFalse);
		//�X���C�h�摜������
		//System::IO::File::Delete(file2);

		//�}�`��
		String^ name;
		//�e�L�X�g�t���[��
		Microsoft::Office::Interop::PowerPoint::TextFrame^ text;
		//�e�L�X�g���e
		String^ text2;
		//�摜�p�X
		String^ picPath;
		//�t�H���g��
		String^ font;
		//�t�H���g�T�C�Y
		int TextSize;
		//�p�l�������ڐ�
		int itemCount = 0;
		//�e�L�X�g���ǂ���
		bool textCheck;
		//�摜���ǂ���
		bool pictureCheck;
		//�\���ǂ���
		bool tableCheck;
		//�}�`�X�^�C��
		MsoShapeStyleIndex style;
		//�}�`�^�C�v
		MsoShapeType type;
		int height;
		int width;
		int x;
		int y;

		
		//picturebox�\���p
		Bitmap^ b;

		//�^�C�g��������΁A�ȉ��̂����Ń^�C�g�������擾�ł���
		if (presense->Slides[1]->Shapes->HasTitle == MsoTriState::msoTrue) {
			String^ titleText = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Text;
			String^ titlefont = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Font->Name;
			String^ titlefont2 = presense->Slides[1]->Shapes->Title->AlternativeText;
			int title_size = presense->Slides[1]->Shapes->Title->TextFrame->TextRange->Font->Size;
		}
		

		

		//presense->Slides[1]->Shapes->Title->Export("C:\\Users\\chach\\Desktop\\sample_title.jpg", PpShapeFormat::ppShapeFormatJPG, width, height, PpExportMode::ppClipRelativeToSlide);

		//�}�`������擾���A����
		for each (Microsoft::Office::Interop::PowerPoint::Shape^ var in presense->Slides[1]->Shapes)
		{
			//�X���C�h�\���p�̃L�����o�X��
			int slideWidth = presense->PageSetup->SlideWidth;
			//�X���C�h�\���p�̃L�����o�X����
			int slideHeight = presense->PageSetup->SlideHeight;
			pictureBox1->Width = slideWidth;
			pictureBox1->Height = slideHeight;
			//�`��p�r�b�g�}�b�v
			b = gcnew Bitmap(pictureBox1->Width, pictureBox1->Height);
			//�^�U�l�ϐ�������
			pictureCheck = false;
			textCheck = false;
			//�}�`�̖��O���擾
			name = var->Name;
			//�}�`�̃^�C�v���擾
			style = var->ShapeStyle;
			type = var->Type;
			//�}�`�̍���
			height = var->Height;
			//�}�`�̕�
			width = var->Width;
			//�}�`�̈ʒu
			x = var->Left;
			y = var->Top;

			

			//�e�L�X�g�^�C�v�̏���
			if (var->HasTextFrame==MsoTriState::msoTrue) {
				textCheck = true;
				//�e�L�X�g���܂܂�Ă���ꍇ
				text = var->TextFrame;
				//�e�L�X�g�擾
				text2 = text->TextRange->Text;
				//�e�L�X�g�T�C�Y
				TextSize = text->TextRange->Font->Size;
				//�e�L�X�g�t�H���g��
				font = var->AlternativeText;
				//var->Export("C:\\Users\\chach\\Desktop\\sample_text"+text+".jpg", PpShapeFormat::ppShapeFormatJPG, width, height, PpExportMode::ppScaleXY);
			}

			if (var->HasTable == MsoTriState::msoTrue) {
				pictureCheck = true;
				//�e�[�u����ۗL���Ă����ꍇ�A��U�͉摜�Ƃ��ĕێ�
				var->Export("C:\\Users\\chach\\Desktop\\sample_hyou.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleXY);
				picPath = "C:\\Users\\chach\\Desktop\\sample_hyou.bmp";
			}
			//���ߍ��݃f�[�^�̏ꍇ
			if (type==MsoShapeType::msoEmbeddedOLEObject) {
				pictureCheck = true;
				var->Export("C:\\Users\\chach\\Desktop\\sample_umekomi.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleToFit);
				picPath = "C:\\Users\\chach\\Desktop\\sample_umekomi.bmp";
			}
			//�摜�t�@�C���̏ꍇ
			if (type == MsoShapeType::msoPicture) {
				pictureCheck = true;
				var->Export("C:\\Users\\chach\\Desktop\\sample_pic.bmp", PpShapeFormat::ppShapeFormatBMP, width, height, PpExportMode::ppScaleXY);
				picPath = "C:\\Users\\chach\\Desktop\\sample_pic.bmp";
			}
			//�}�`�\���̂ɑ��
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
			//�}�`�����擾
			shapeList->Add(var);
			shapesList->Add(temp);
		}
		int count = 0;
		//�擾�����}�`�����p�l����pictureBox�ɕ\���������Ă���
		for each (Ashapes^ var in shapesList)
		{
			//�摜�̏ꍇ
			if (var->picture) {
				//�p�l���\���p�̃s�N�`���[�{�b�N�X
				pic = gcnew PictureBox;
				//�R���g���[���ɒǉ�
				this->Controls->Add(pic);
				//�s�N�`���[�{�b�N�X�̍��W
				pic->Top = 25 * itemCount;
				//����
				pic->Height = 25;
				//�摜�p�X����C���[�W�쐬
				System::Drawing::Image^ img= System::Drawing::Image::FromFile(var->picturePath);
				//�C���[�W�̂͂ߍ���
				pic->Image = img;
				//�\���p�̃s�N�`���[�{�b�N�X�̃C���[�W����̃r�b�g�}�b�v�ɐݒ�
				pictureBox1->Image = b;
				//�O���t�B�b�N��ݒ�
				Graphics^ gr = Graphics::FromImage(pictureBox1->Image);
				//�摜�\��
				gr->DrawImage(img, var->x, var->y, var->width, var->height);
				//�`��X�V
				pictureBox1->Refresh();
				count++;
				pic->MouseDown += gcnew System::Windows::Forms::MouseEventHandler(this, &MyForm::picture_MouseDown);
				
			}
			else {
				//���x���쐬
				tx = gcnew Label;
				//�R���g���[���ɒǉ�
				this->Controls->Add(tx);
				tx->Top = 25 * itemCount;
				tx->Text = var->textVal;
				//�s�N�`���[�{�b�N�X�쐬
				//��̃r�b�g�}�b�v�ɕ����`��
				//�t�H���g�Ȃǎw��
				tx->MouseDown += gcnew System::Windows::Forms::MouseEventHandler(this, &MyForm::text_MouseDown);
			}
			itemCount++;
		}
		//�p���[�|�C���g�̃C���X�^���X����
		presense->Close();
		app_->Quit();

	}
	return System::Void();
}

System::Void samplepp::MyForm::MyForm_DragEnter(System::Object ^ sender, System::Windows::Forms::DragEventArgs ^ e)
{
	//�t�H�[���ւ̃t�@�C���h���b�O�h���b�v�C�x���g
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
	//�p�l������̃h���b�O�h���b�v�C�x���g
	if (e->Data->GetDataPresent(DataFormats::Bitmap)) {
		System::Drawing::Graphics^ tegr = pictureBox1->CreateGraphics();
		Image^ img = (Image^)e->Data->GetData(DataFormats::Bitmap);
		int xp = pictureBox1->MousePosition.X;
		int yp = pictureBox1->MousePosition.Y;
		tegr->DrawImage((Image^)e->Data->GetData(DataFormats::Bitmap), xp, yp);
		//pictureBox1->Image = (Image^)e->Data->GetData(DataFormats::Bitmap);

	}
	else if (e->Data->GetDataPresent(DataFormats::FileDrop)) {
		//�f�[�^�I�u�W�F�N�g�̎󂯎��
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
	//�}�E�X�N���b�N�����s�N�`���[�{�b�N�X���擾
	pic = (PictureBox^)sender;
	//�h���b�O�h���b�v�C�x���g��ݒ�
	pic->DoDragDrop(pic->Image, DragDropEffects::All);
	return System::Void();
}

System::Void samplepp::MyForm::MyForm_Load(System::Object ^ sender, System::EventArgs ^ e)
{
	//���[�h���A�h���b�v�C�x���g������
	pictureBox1->AllowDrop = true;
	//�C�x���g�쓮��ݒ�
	pictureBox1->DragDrop += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::picture_DragDrop);
	pictureBox1->DragEnter += gcnew System::Windows::Forms::DragEventHandler(this, &MyForm::picture_DragDrop);
	return System::Void();
}


System::Void samplepp::MyForm::text_MouseDown(System::Object ^ sender, System::Windows::Forms::MouseEventArgs ^ e)
{
	//�p�l�����烉�x����I�������ꍇ�̏���
	text_pickUp = true;
	//�K�v�ȃf�[�^�̔z����쐬
	array<String^>^ data = gcnew array<String^>(3);
	//�ǂ̃��x�����I�����ꂽ��
	tx = (Label^)sender;
	//���x���̃e�L�X�g���擾
	String^ lbtext = tx->Text;
	//�e�L�X�g�ƈ�v����}�`�����������A�e�L�X�g���e�A�t�H���g�A�T�C�Y���擾���A�z��Ɋi�[
	for each (Ashapes^ var in shapesList)
	{
		if (var->text&&(lbtext == var->textVal)) {
			data[0] = var->textVal;
			data[1] = var->fontName;
			data[2] = var->textSize.ToString();
		}
	}
	//�f�[�^�I�u�W�F�N�g�Ƃ��Ă܂Ƃ߁A�t�@�C���h���b�v�`���ŕR�Â���
	DataObject^ dobj = gcnew DataObject(DataFormats::FileDrop, data);
	//�h���b�O�h���b�v�C�x���g�Ńf�[�^�I�u�W�F�N�g����������
	tx->DoDragDrop(dobj, DragDropEffects::All);
	return System::Void();
}


#pragma once

namespace ЛБМаксимюк9 {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;
	using namespace ADOX;
	using namespace Microsoft::Office::Interop::Access;
	using namespace System::Data::OleDb;

	/// <summary>
	/// Сводка для MyForm
	/// </summary>
	public ref class MyForm : public System::Windows::Forms::Form
	{
	public:
		MyForm(void)
		{
			InitializeComponent();
			//
			//TODO: добавьте код конструктора
			//
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~MyForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ Struct;
	private: System::Windows::Forms::Button^ Read;
	private: System::Windows::Forms::DataGridView^ dataGridView1;
	private: System::Windows::Forms::Button^ button1;
	protected:


	protected:

	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->Struct = (gcnew System::Windows::Forms::Button());
			this->Read = (gcnew System::Windows::Forms::Button());
			this->dataGridView1 = (gcnew System::Windows::Forms::DataGridView());
			this->button1 = (gcnew System::Windows::Forms::Button());
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->BeginInit();
			this->SuspendLayout();
			// 
			// Struct
			// 
			this->Struct->Anchor = System::Windows::Forms::AnchorStyles::Bottom;
			this->Struct->BackColor = System::Drawing::SystemColors::ButtonFace;
			this->Struct->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->Struct->Location = System::Drawing::Point(44, 410);
			this->Struct->Name = L"Struct";
			this->Struct->Size = System::Drawing::Size(135, 42);
			this->Struct->TabIndex = 0;
			this->Struct->Text = L"Структура";
			this->Struct->UseVisualStyleBackColor = false;
			this->Struct->Click += gcnew System::EventHandler(this, &MyForm::Struct_Click);
			// 
			// Read
			// 
			this->Read->Anchor = System::Windows::Forms::AnchorStyles::Bottom;
			this->Read->BackColor = System::Drawing::SystemColors::ButtonFace;
			this->Read->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->Read->Location = System::Drawing::Point(214, 410);
			this->Read->Name = L"Read";
			this->Read->Size = System::Drawing::Size(135, 42);
			this->Read->TabIndex = 1;
			this->Read->Text = L"Зчитування";
			this->Read->UseVisualStyleBackColor = false;
			this->Read->Click += gcnew System::EventHandler(this, &MyForm::Read_Click);
			// 
			// dataGridView1
			// 
			this->dataGridView1->Anchor = static_cast<System::Windows::Forms::AnchorStyles>((((System::Windows::Forms::AnchorStyles::Top | System::Windows::Forms::AnchorStyles::Bottom)
				| System::Windows::Forms::AnchorStyles::Left)
				| System::Windows::Forms::AnchorStyles::Right));
			this->dataGridView1->AutoSizeColumnsMode = System::Windows::Forms::DataGridViewAutoSizeColumnsMode::Fill;
			this->dataGridView1->AutoSizeRowsMode = System::Windows::Forms::DataGridViewAutoSizeRowsMode::DisplayedCells;
			this->dataGridView1->ColumnHeadersHeightSizeMode = System::Windows::Forms::DataGridViewColumnHeadersHeightSizeMode::AutoSize;
			this->dataGridView1->Location = System::Drawing::Point(12, 12);
			this->dataGridView1->Name = L"dataGridView1";
			this->dataGridView1->Size = System::Drawing::Size(372, 392);
			this->dataGridView1->TabIndex = 2;
			// 
			// button1
			// 
			this->button1->Anchor = System::Windows::Forms::AnchorStyles::Bottom;
			this->button1->Location = System::Drawing::Point(149, 463);
			this->button1->Name = L"button1";
			this->button1->Size = System::Drawing::Size(96, 23);
			this->button1->TabIndex = 3;
			this->button1->Text = L" Створити БД";
			this->button1->UseVisualStyleBackColor = true;
			this->button1->Click += gcnew System::EventHandler(this, &MyForm::button1_Click);
			// 
			// MyForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->BackColor = System::Drawing::SystemColors::ActiveCaption;
			this->ClientSize = System::Drawing::Size(396, 498);
			this->Controls->Add(this->button1);
			this->Controls->Add(this->dataGridView1);
			this->Controls->Add(this->Read);
			this->Controls->Add(this->Struct);
			this->Name = L"MyForm";
			this->Text = L"База Даних";
			(cli::safe_cast<System::ComponentModel::ISupportInitialize^>(this->dataGridView1))->EndInit();
			this->ResumeLayout(false);

		}
#pragma endregion

		DataSet^ Dataset;
		OleDb::OleDbDataAdapter^ Adapter;
		OleDb::OleDbConnection^ Connections;
		OleDb::OleDbCommand^ Commands;
	private: System::Void Struct_Click(System::Object^ sender, System::EventArgs^ e) {
		OleDbConnection^ Connect = gcnew OleDbConnection("Provider=Microsoft.Jet." + "OLEDB.4.0;Data Source=D:\\New_BD.mdb");
		Connect->Open();
		OleDbCommand^ Command = gcnew OleDbCommand("CREATE TABLE [DB PhoneNumbers]"+"([Номер п/п] counter , [ПІП] char(50),"+"[Номер телефону] char(15))", Connect);
		try
		{
			Command->ExecuteNonQuery();
			MessageBox::Show("Структура таблиці 'DB PhoneNumbers' записана в порожню БД", "Створення структури таблиці MS Access", MessageBoxButtons::OK, MessageBoxIcon::Information);
		}
		catch (Exception^ Situation)
		{
			MessageBox::Show(Situation->Message, "Створення структури таблиці MS Access", MessageBoxButtons::OK, MessageBoxIcon::Warning);
		}
		Connect->Close();
	}
private: System::Void Read_Click(System::Object^ sender, System::EventArgs^ e) {
	auto Connections = gcnew OleDb::OleDbConnection("Data Source=D:\\New_BD.mdb; User ID="+"Admin; Provider = Microsoft.Jet." + "OLEDB.4.0;");
	Connections->Open();
	auto Commands = gcnew OleDb::OleDbCommand("Select * From [DB PhoneNumbers]", Connections);
	auto Adapter = gcnew OleDb::OleDbDataAdapter(Commands);
	auto Dataset = gcnew DataSet();
	Adapter->Fill(Dataset, "DB PhoneNumbers");
	auto РядокXML = Dataset->GetXml();
	dataGridView1->DataSource = Dataset;
	dataGridView1->DataMember = "DB PhoneNumbers";
	Connections->Close();
}
private: System::Void button1_Click(System::Object^ sender, System::EventArgs^ e) {
	ADOX::Catalog^ Catalog = gcnew ADOX::Catalog();
	try
	{
		Catalog->Create("Provider=Microsoft.Jet." + "OLEDB.4.0;Data Source=D:\\New_BD.mdb");
		MessageBox::Show("База Даних D:\\New_BD.mdb створена!", "Створення нової БД MS Access", MessageBoxButtons::OK, MessageBoxIcon::Information);
	}
	catch (System::Runtime::InteropServices::COMException^ Situation)
	{
		MessageBox::Show(Situation->Message, "База Даних вже створена!", MessageBoxButtons::OK, MessageBoxIcon::Warning);
	}
	finally {
		Catalog = nullptr;
	}
}
};
}

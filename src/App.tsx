import './App.css';
import FileSaver from "file-saver";
import { Document, Packer, Paragraph, TextRun } from "docx";

function App() {
  function generateDoc(){
    console.log('teste');
    const doc = new Document({
      sections: [
          {
              properties: {},
              children: [
                  new Paragraph({
                      children: [
                          new TextRun({text: 'Novo Relatorio', size: 12}),
                      ],
                  }),
                //   new Paragraph({
                //     children: [
                //         new TextRun({
                //             text: "Foo Bar",
                //             bold: true,
                //         }),
                //         new TextRun({
                //             text: "\nNovo teste",
                //             bold: true,
                //         }),
                //     ],
                // }),
              ],
          },
      ],
  });

  Packer.toBlob(doc).then((blob) => {
    FileSaver.saveAs(blob, "example.docx")
});
  }

  return (
    <div className="App">
      <header className="App-header">
        <button onClick={generateDoc}>clique aqui</button>
      </header>
    </div>
  );
}

export default App;

<!DOCTYPE html>
<html>
<head>
  <title>Web Scraping and Text Analysis Tool</title>
</head>
<body>
  <h1>Web Scraping and Text Analysis Tool</h1>

  <p>This Python script is a web scraping and text analysis tool that extracts data from a list of URLs provided in an Excel file. It performs various text analysis tasks and saves the results to an output Excel file. The extracted articles are also saved as individual text files.</p>

  <h2>Getting Started</h2>
  <p>Before using this tool, you need to install the required Python libraries. You can install them using pip:</p>

  <pre>
    <code>pip install requests
pip install beautifulsoup4
pip install openpyxl
pip install textblob
pip install textstat</code>
  </pre>

  <h2>Usage</h2>
  <ol>
    <li>Clone or download this repository to your local machine.</li>
    <li>Open the <code>Input.xlsx</code> file and add the URLs you want to analyze. Make sure the Excel file has the following structure:</li>
  </ol>

  <table>
    <tr>
      <th>URL_ID</th>
      <th>URL</th>
    </tr>
    <tr>
      <td>1</td>
      <td>https://example.com/article1</td>
    </tr>
    <tr>
      <td>2</td>
      <td>https://example.com/article2</td>
    </tr>
    <!-- ... -->
  </table>

  <p>3. Update the <code>input_file_path</code> and <code>output_file_path</code> variables in the Python script to specify the paths to your input and output Excel files.</p>

  <p>4. Run the Python script. It will extract data from the URLs, perform text analysis, and save the results to the output Excel file.</p>

  <h2>Output</h2>
  <p>The tool will generate an output Excel file with the following columns:</p>

  <ul>
    <li>URL_ID</li>
    <li>Sentiment_Polarity</li>
    <li>Sentiment_Subjectivity</li>
    <li>FOG_Index (Readability)</li>
    <li>Complex_Word_Count</li>
    <li>Word_Count</li>
    <li>Syllables_Per_Word</li>
    <li>Personal_Pronouns</li>
    <li>Avg_Word_Length</li>
    <li>Avg_Sentence_Length</li>
    <li>Avg_Words_Per_Sentence</li>
  </ul>

  <p>In addition to the Excel file, the extracted articles will be saved in a directory called <code>Extracted_Articles</code>. Each article is saved as a separate text file with the format <code>URL_ID.txt</code>.</p>

  <h2>Customization</h2>
  <p>You can customize the tool by adding more text analysis functions or modifying the existing ones according to your needs. For example, you can add functions to extract specific information from the web pages.</p>

  <h2>Issues and Contributions</h2>
  <p>If you encounter any issues or have suggestions for improvements, please create a GitHub issue or submit a pull request.</p>

  <hr>

  <p>This tool provides a convenient way to scrape web content, perform text analysis, and store the results for further analysis or reporting. It can be particularly useful for researchers, content analysts, or anyone interested in automating the extraction and analysis of text data from websites.</p>
</body>
</html>

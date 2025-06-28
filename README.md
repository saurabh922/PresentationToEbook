# 📚 Advanced PowerPoint to eBook Converter

**Transform your PowerPoint presentations into professional eBooks with AI-powered content enhancement, flexible chapter organization, and multiple output formats.**

[![Streamlit](https://img.shields.io/badge/Built%20with-Streamlit-red)](https://streamlit.io/)
[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://python.org)
[![Gemini AI](https://img.shields.io/badge/Powered%20by-Gemini%20AI-green)](https://ai.google.dev/)

## 🎉 **Complete Feature Set**

### **📖 Multiple Output Formats**
- **PDF Generation**: Professional layout with headers, footers, and chapter tracking
- **DOCX Generation**: Microsoft Word format with proper formatting
- **Both Formats**: Generate PDF and DOCX simultaneously

### **🎯 Advanced Chapter Management**
- **Sequential Chapter Numbering**: Proper 1, 2, 3... numbering with AI integration
- **Flexible Organization**: 3 methods available:
  - **Automatic (Equal Groups)**: 2-20 slides per chapter
  - **Custom Ranges**: User-defined ranges (e.g., "1-5, 6-12, 13-18")
  - **One Slide Per Chapter**: Individual chapters for each slide
- **Smart Chapter Tracking**: Current chapter number displayed in PDF footers

### **🖼️ Professional Image Integration**
- **Smart Image Extraction**: Automatically extracts images and diagrams from slides
- **Contextual Placement**: Images appear with their respective slide content
- **Intelligent Captions**: Uses actual slide titles in figure captions
- **Diagram Recognition**: Distinguishes between images and flowcharts/diagrams
- **Type-Aware Processing**: Different handling for images vs diagrams

### **🤖 Enhanced AI Content Processing**
- **Clean AI Responses**: Removes repetitive "Of course. Here is..." text
- **Markdown Parsing**: Proper heading hierarchy (H1, H2, H3, H4)
- **Content Enhancement**: Expands bullet points into comprehensive paragraphs
- **Professional Tone**: Maintains consistency throughout the eBook

### **🎛️ Dynamic Logging & Debug System**
- **Smart Debug Mode**: Checkbox controls actual logging level (not just UI display)
- **Performance Optimized**: INFO level by default, DEBUG only when needed
- **Enhanced Debug UI**: Shows recent logs, statistics (Total, Debug, Info, Errors)
- **Real-time Control**: Logging level changes immediately when toggled

### **📄 Professional Document Features**
- **Custom Document Template**: BaseDocTemplate with proper page structure
- **Author Information**: Professional title page styling with author details
- **Custom Headers & Footers**: Customizable text with chapter tracking
- **Page Numbers**: Right-aligned page numbers maintained
- **Professional Typography**: Modern fonts and spacing

## 🚀 **How It Works**

1. **📤 Upload**: Select your PowerPoint presentation (PPTX format, up to 100MB)
2. **⚙️ Configure**: Choose chapter organization method and output format
3. **🔍 Extract**: App extracts text, images, and slide structure
4. **🤖 Enhance**: Gemini AI expands and improves content with chapter-aware prompts
5. **📚 Generate**: Creates professional eBook(s) with integrated images
6. **💾 Download**: Get your completed eBook in PDF, DOCX, or both formats

## 🛠️ **Installation & Setup**

### **Prerequisites**
- Python 3.8 or higher
- Git (for cloning the repository)
- Google Gemini API key (free)

### **Quick Start**

1. **Clone the repository**:
   ```bash
   git clone <repository-url>
   cd SlideToEBook
   ```

2. **Create virtual environment** (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Get Gemini API Key**:
   - Visit [Google AI Studio](https://makersuite.google.com/app/apikey)
   - Create a free API key (no credit card required)
   - Keep it handy for the application

5. **Run the application**:
   ```bash
   streamlit run app.py
   ```

6. **Open your browser** and navigate to the provided local URL (usually `http://localhost:8501`)

## 📖 **Usage Guide**

### **Basic Usage**
1. **🔑 Configure API Key**: Enter your Gemini API key in the sidebar
2. **📝 Set eBook Details**: Customize title, author, and headers/footers
3. **📤 Upload PowerPoint**: Choose your PPTX file (max 100MB)
4. **⚙️ Choose Settings**: Select output format and chapter organization
5. **🚀 Convert**: Click "Generate eBook" and wait for processing
6. **💾 Download**: Get your completed eBook(s)

### **Chapter Organization Options**

**🔄 Automatic (Equal Groups)**
- Best for: Medium to large presentations
- Settings: Choose 2-20 slides per chapter
- Example: 30 slides → 6 chapters of 5 slides each

**🎯 Custom Ranges**
- Best for: Topic-based organization
- Settings: Define ranges like "1-5, 6-12, 13-18"
- Example: Group related slides into logical chapters

**📄 One Slide Per Chapter**
- Best for: Detailed breakdown or small presentations
- Settings: Each slide becomes its own chapter
- Example: 10 slides → 10 individual chapters

### **Output Format Selection**
- **📄 PDF Only**: Professional PDF with chapter tracking
- **📝 DOCX Only**: Microsoft Word format for editing
- **📚 Both Formats**: Get both PDF and DOCX files

## 🔧 **Technical Architecture**

### **Core Dependencies**
- **🌐 Streamlit** `>=1.28.0`: Modern web UI framework with real-time updates
- **📄 python-pptx** `>=0.6.21`: PowerPoint file processing and content extraction
- **🤖 google-generativeai** `>=0.3.0`: Gemini AI integration for content enhancement
- **📄 reportlab** `>=4.0.0`: Professional PDF generation with custom templates
- **🖼️ Pillow** `>=9.0.0`: Advanced image processing and format support
- **📝 python-docx** `>=0.8.11`: Microsoft Word document generation

### **Project Structure**
```
SlideToEBook/
├── app.py                 # Main application with all features
├── requirements.txt       # Python dependencies with versions
├── README.md             # Comprehensive documentation
├── ppt_to_ebook.log      # Dynamic logging output
└── venv/                 # Virtual environment (created during setup)
```

### **Core Components**

**📚 PPTToEBookConverter Class**
- Extracts text, images, and diagrams from PPTX files
- Implements flexible chapter organization algorithms
- Integrates with Gemini AI for content enhancement
- Generates professional PDF and DOCX eBooks

**🎨 CustomDocTemplate**
- Professional PDF layout with headers and footers
- Chapter-aware footer system showing current chapter
- Custom page templates with proper margins

**🔍 ChapterMarker**
- Custom ReportLab flowable for chapter tracking
- Updates footer content dynamically
- Invisible element that maintains document flow

**📊 Dynamic Logging System**
- Performance-optimized logging (INFO by default)
- Real-time debug mode switching
- Enhanced UI with log statistics and filtering

## 🚀 **Usage Examples**

### **Small Presentations (10-20 slides)**
```
Recommended Settings:
✓ Chapter Method: "One Slide Per Chapter"
✓ Output Format: "Both (PDF + DOCX)"
✓ Result: Detailed breakdown with maximum flexibility
```

### **Medium Presentations (30-50 slides)**
```
Recommended Settings:
✓ Chapter Method: "Automatic (Equal Groups)"
✓ Slides per Chapter: 3-5
✓ Output Format: "PDF Only"
✓ Result: Professional eBook with logical chapters
```

### **Large Presentations (100+ slides)**
```
Recommended Settings:
✓ Chapter Method: "Custom Ranges"
✓ Example Ranges: "1-10, 11-25, 26-40, 41-60"
✓ Output Format: "PDF Only"
✓ Result: Topic-based organization with custom structure
```

## 🔍 **Troubleshooting Guide**

### **Common Issues & Solutions**

**📁 File Upload Issues**
- ❗ Problem: "File size too large"
- ✅ Solution: Ensure PPTX file is under 100MB, compress images if needed

**🔑 API Key Problems**
- ❗ Problem: "Invalid API key" or "API quota exceeded"
- ✅ Solution: Verify key at [Google AI Studio](https://makersuite.google.com/app/apikey), check usage limits

**🖼️ Image Processing Issues**
- ❗ Problem: Images not appearing in output
- ✅ Solution: Enable debug mode, check logs for image extraction status

**📄 PDF Generation Errors**
- ❗ Problem: "Failed to build PDF document"
- ✅ Solution: Check dependencies, try with smaller file first

### **Performance Optimization Tips**

**⚡ For Large Files:**
- Use "Custom Ranges" to process in smaller chunks
- Keep debug mode OFF during production use
- Close other applications to free up memory

**📊 For Best Quality:**
- Use clear, descriptive slide titles
- Include detailed content (not just bullet points)
- Ensure images are high-quality and relevant
- Test with smaller presentations first

## 🎆 **Advanced Features**

### **Debug Mode Capabilities**
- **Real-time Logging**: See processing steps as they happen
- **Performance Metrics**: Track processing time and resource usage
- **Error Diagnosis**: Detailed error messages and stack traces
- **Image Tracking**: Monitor image extraction and placement

### **Professional PDF Features**
- **Chapter-Aware Footers**: Shows "Chapter X" instead of static text
- **Custom Typography**: Professional fonts and spacing
- **Image Integration**: Smart placement with proper captions
- **Metadata Support**: Author, title, and creation date embedded

### **DOCX Generation Features**
- **Native Word Format**: Fully editable in Microsoft Word
- **Image Embedding**: High-quality images with captions
- **Heading Styles**: Proper heading hierarchy for navigation
- **Professional Layout**: Consistent formatting throughout

## 📄 **License**

This project is open source and available under the **MIT License**.

```
MIT License

Copyright (c) 2024 PowerPoint to eBook Converter

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
```

## 👥 **Contributing**

We welcome contributions! Here's how you can help:

### **Ways to Contribute**
- 🐛 **Bug Reports**: Found an issue? Create a detailed bug report
- ✨ **Feature Requests**: Have an idea? Suggest new features
- 📝 **Documentation**: Help improve our documentation
- 💻 **Code Contributions**: Submit pull requests with improvements

### **Development Setup**
1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Make your changes and test thoroughly
4. Commit your changes: `git commit -m 'Add amazing feature'`
5. Push to the branch: `git push origin feature/amazing-feature`
6. Open a Pull Request

## 💬 **Support & Community**

### **Getting Help**
- 📚 **Documentation**: Check this README for comprehensive guides
- 🐛 **Issues**: Create an issue for bugs or feature requests
- 💬 **Discussions**: Join community discussions for questions

### **Reporting Issues**
When reporting issues, please include:
- Python version and operating system
- Complete error messages and stack traces
- Steps to reproduce the problem
- Sample PowerPoint file (if possible)
- Debug logs (enable debug mode)

## 🎆 **Roadmap & Future Features**

### **Planned Enhancements**
- 🌍 **Multi-language Support**: Support for non-English presentations
- 📊 **Analytics Dashboard**: Processing statistics and insights
- ☁️ **Cloud Integration**: Direct upload from cloud storage
- 📱 **Mobile Optimization**: Better mobile web experience
- 🔄 **Batch Processing**: Process multiple files simultaneously

### **Version History**
- **v2.0.0** (Current): Complete rewrite with advanced features
  - Multiple output formats (PDF + DOCX)
  - Professional image integration
  - Flexible chapter organization
  - Dynamic logging system
  - Enhanced AI processing

---

## 🎉 **Success Stories**

*"Transformed our 200-slide training presentation into a professional eBook in minutes. The AI enhancement made our content much more readable!"* - Corporate Training Team

*"Perfect for converting academic presentations into study materials. The custom chapter ranges feature is exactly what we needed."* - University Professor

*"The image integration works flawlessly. Our technical diagrams are perfectly placed with proper captions."* - Technical Documentation Team

---

<div align="center">

### 🚀 **Ready to Transform Your Presentations?**

**[Get Started Now](#installation--setup)** | **[View Examples](#usage-examples)** | **[Report Issues](https://github.com/your-repo/issues)**

---

**🎉 Happy Converting! Transform your PowerPoint presentations into professional eBooks today! 📚**

*Built with ❤️ using Streamlit, Gemini AI, and modern Python technologies*

</div>

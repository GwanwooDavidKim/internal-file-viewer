# Overview

The Internal File Viewer (사내 파일 뷰어) is a desktop application designed to help teams efficiently browse and preview various business document formats (PDF, PowerPoint, Excel, Word, images) within a single unified interface. The application eliminates the need to open multiple applications to view different file types, significantly improving workflow efficiency for teams dealing with diverse document collections.

The system features a secure authentication system with role-based access control, comprehensive file format support with preview capabilities, and a powerful full-text search engine with indexing. Built as a desktop application using PyQt6, it operates completely offline without requiring external network connections.

# User Preferences

Preferred communication style: Simple, everyday language.

# System Architecture

## Desktop Application Framework
The application is built using PyQt6 as the primary GUI framework, providing a native desktop experience with a modern, intuitive interface. The architecture follows a modular design pattern with clear separation of concerns across different functional areas.

## User Interface Design
The UI implements a 2-pane layout consisting of a file explorer on the left and a content viewer on the right. The interface includes a tabbed system supporting both file browsing and search functionality. The design emphasizes clarity and efficiency, using a professional color palette with dark greys and blues to reduce eye strain during extended use.

## Authentication and Security System
The application implements a robust authentication system using bcrypt for password hashing. It supports two user roles: administrators and regular team members. The system includes account expiration management where administrators can set usage periods for team member accounts. All authentication data is stored locally in the configuration file with hashed passwords.

## File Processing Architecture
The system uses a modular file handler architecture where each file type has a dedicated handler class. A central FileManager class coordinates between different handlers and provides a unified interface. The architecture supports dynamic file type detection and extensible handler registration.

## Multi-threaded Processing
File loading and search indexing operations are performed in background threads to maintain UI responsiveness. The system uses PyQt6's QThread for worker threads, with proper signal-slot communication for progress updates and result delivery.

## Search and Indexing System
The application features a sophisticated full-text search engine with inverted index architecture. The SearchIndexer builds and maintains indices of file contents, supporting both Korean and English text with stopword filtering. The system provides real-time search with result highlighting and context preview.

## Data Storage and Caching
The application uses in-memory data structures for fast access to indexed content and file metadata. Search indices are built dynamically and can be persisted for performance optimization. The system implements caching strategies for frequently accessed file previews and metadata.

## Error Handling and Logging
Comprehensive error handling is implemented throughout the application with user-friendly error messages displayed through the GUI. A structured logging system captures application events, file operations, authentication attempts, and performance metrics with automatic log rotation.

# External Dependencies

## GUI Framework
- **PyQt6**: Primary desktop GUI framework providing native widget support and event handling
- **PyQt6-Qt6**: Core Qt6 libraries
- **PyQt6-sip**: Python bindings for Qt6

## File Processing Libraries
- **PyMuPDF**: PDF file processing for page rendering and text extraction
- **python-pptx**: PowerPoint file text extraction and structure analysis
- **python-docx**: Word document text extraction and metadata processing
- **Pillow**: Image file processing and display with support for multiple formats
- **pandas**: Excel file data reading and manipulation
- **openpyxl**: Excel file reading/writing with metadata support
- **xlsxwriter**: Enhanced Excel file support for pandas operations

## Security and Authentication
- **bcrypt**: Password hashing and verification for secure authentication

## Data Processing
- **numpy**: Numerical operations support (dependency of pandas)
- **lxml**: XML processing for document parsing
- **et-xmlfile**: XML file handling for Excel operations

## Utility Libraries
- **python-dateutil**: Date and time processing utilities
- **pytz**: Timezone handling
- **tzdata**: Timezone data
- **typing-extensions**: Enhanced type hinting support
- **six**: Python 2/3 compatibility utilities

## Runtime Dependencies
The application is designed to be packaged as a standalone executable using PyInstaller, eliminating the need for users to install Python or manage dependencies. For PowerPoint slide rendering (if COM automation is implemented), Microsoft PowerPoint installation on the user's machine would be required.
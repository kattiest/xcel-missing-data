#include <iostream>
#include <vector>
#include "libxl.h"

using namespace libxl;

int main(int argc, char* argv[]) {
    if (argc < 4) {
        std::cout << "Usage: " << argv[0] << " <file1.xls> <file2.xls> ... <column_number>" << std::endl;
        return 1;
    }

    // Get the column number to check for duplicates
    int columnNumber = std::stoi(argv[argc - 1]);

    // Load Excel files
    Book* book;
    std::vector<Sheet*> sheets;
    for (int i = 1; i < argc - 1; i++) {
        book = xlCreateXMLBook();
        if (book->load(argv[i])) {
            Sheet* sheet = book->getSheet(0);
            sheets.push_back(sheet);
        } else {
            std::cout << "Failed to load file: " << argv[i] << std::endl;
            return 1;
        }
    }

    // Find duplicates in the specified column
    std::vector<std::string> duplicates;
    for (Sheet* sheet : sheets) {
        int rowCount = sheet->lastRow() + 1;
        for (int row = 0; row < rowCount; row++) {
            const std::string cellValue = sheet->readStr(row, columnNumber);
            for (int r = row + 1; r < rowCount; r++) {
                const std::string compareValue = sheet->readStr(r, columnNumber);
                if (cellValue == compareValue && cellValue != "") {
                    duplicates.push_back(cellValue);
                }
            }
        }
    }

    // Print duplicates
    if (!duplicates.empty()) {
        std::cout << "Duplicates found in column " << columnNumber << ":" << std::endl;
        for (const std::string& duplicate : duplicates) {
            std::cout << duplicate << std::endl;
        }
    } else {
        std::cout << "No duplicates found." << std::endl;
    }

    // Cleanup
    for (Sheet* sheet : sheets) {
        delete sheet;
    }
    delete book;

    return 0;
}

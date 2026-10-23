#include "MarkdownParser.h"
#include <sstream>
#include <cctype>

namespace ppt {

static std::string trim(const std::string& s) {
  size_t start = 0, end = s.size();
  while (start < end && std::isspace(static_cast<unsigned char>(s[start]))) ++start;
  while (end > start && std::isspace(static_cast<unsigned char>(s[end - 1]))) --end;
  return s.substr(start, end - start);
}

Document MarkdownParser::parse(const std::string& text) const {
  Document doc;
  Slide current;
  bool hasCurrent = false;

  std::istringstream iss(text);
  std::string line;

  while (std::getline(iss, line)) {
    line = trim(line);
    if (line.empty()) continue;

    if (line == "---") {
      if (hasCurrent) {
        doc.slides.push_back(current);
        current = Slide{};
        hasCurrent = false;
      }
      continue;
    }

    if (line[0] == '#') {
      if (hasCurrent) doc.slides.push_back(current);
      size_t level = 0;
      while (level < line.size() && line[level] == '#') ++level;
      current = Slide{};
      current.title = trim(line.substr(level));
      current.layout = "cover";
      hasCurrent = true;
    } else {
      if (!hasCurrent) {
        current = Slide{};
        current.title = "未命名";
        current.layout = "cover";
        hasCurrent = true;
      }
      std::string item = line;
      if (item[0] == '-' || item[0] == '*' || item[0] == '+') item = trim(item.substr(1));
      else {
        size_t i = 0;
        while (i < item.size() && std::isdigit(static_cast<unsigned char>(item[i]))) ++i; // 判断是否为带序号的输入
        if (i > 0 && i < item.size() && item[i] == '.') item = trim(item.substr(i + 1)); // 确定数字不是列表序号
      }
      current.bullets.push_back(item);
    }
  }
  if (hasCurrent) doc.slides.push_back(current); // 如果当前有处于构建中的slide，将其压入doc中

  if (!doc.slides.empty()) doc.title = doc.slides[0].title;
  return doc;
}

}

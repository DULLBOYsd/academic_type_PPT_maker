#pragma once
#include "../model/Document.h"
#include <string> 

namespace ppt {

class MarkdownParser {
public:
  Document parse(const std::string& text) const;
};

}

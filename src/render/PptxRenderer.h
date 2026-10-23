#pragma once 
#include "../model/Document.h"
#include <string> 

namespace ppt {
class PptxRenderer {
public: 
  // 将Document渲染为pptx文件，返回是否成功
  bool render(const Document& doc, const std::string& outputPath);
};

}

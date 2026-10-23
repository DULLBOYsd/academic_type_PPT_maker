#include "io/FileReader.h"
#include "parser/MarkdownParser.h"
#include "render/PptxRenderer.h"

// JSON文件输出
#include <nlohmann/json.hpp>
#include <fstream>

#include <iostream>

// 将Documnet转换为JSON并保存到文件
bool saveAsJson(const ppt::Document& doc, const std::string& path) {
  nlohmann::json j;
  j["title"] = doc.title;

  // 遍历每一页
  for (const auto& slide : doc.slides) {
    nlohmann::json slideJson;
    slideJson["title"] = slide.title;
    slideJson["bullets"] = slide.bullets;
    slideJson["layout"] = slide.layout;
    slideJson["conclusion"] = slide.conclusion;
    j["slides"].push_back(slideJson);
  }

  // 写入文件（注意使用std::ios::binary避免Windows下的换行符问题
  std::ofstream ofs(path, std::ios::binary);
  if (!ofs) {
    std::cerr << "无法写入文件: " << path << "\n";
    return false;
  }

  // dump(4)表示用4个空格缩进，方便人阅读
  ofs << j.dump(4);
  return true;
}

int main(int argc, char* argv[]) {
  if (argc < 3) {
    std::cerr << "用法: " << argv[0] << " input.md output.pptx\n";
    return 1;
  }

  try {
    std::string text = ppt::readFile(argv[1]);
    ppt::MarkdownParser parser;
    ppt::Document doc = parser.parse(text);
    std::cout << "解析完成，共 " << doc.slides.size() << " 页\n";

    std::string outputPath = argv[2];

    // 如果输出文件以.json结尾，就导出JSON数据
    if (outputPath.size() > 5 && outputPath.substr(outputPath.size() - 5) == ".json") {
      if (saveAsJson(doc, outputPath)) std::cout <<  "已生成JSON数据: " << outputPath << "\n";
      else {
        std::cerr << "JSON导出失败\n";
        return 1;
      }
    }
    // 否则，采用原来的PPTX渲染流程
    else {
      ppt::PptxRenderer renderer;
      if (!renderer.render(doc, argv[2])) { // 只需要检测第2帧是否能被渲染即可
        std::cerr << "渲染失败\n";
        return 1;
      }
      std::cout << "已生成: " << argv[2] << "\n";
    }
  } catch (const std::exception& e) {
    std::cerr << "错误: " << e.what() << "\n";
    return 1;
  }
  return 0;
}

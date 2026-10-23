#include "PptxRenderer.h"
#include <Aspose/Slides/Foss/presentation.h>
#include <Aspose/Slides/Foss/slide.h>
#include <Aspose/Slides/Foss/shape_collection.h>
#include <Aspose/Slides/Foss/auto_shape.h>
#include <Aspose/Slides/Foss/shape_type.h>
#include <Aspose/Slides/Foss/text_frame.h>
#include <Aspose/Slides/Foss/export/save_format.h>
#include <iostream>

using namespace Aspose::Slides::Foss;

namespace ppt {

bool PptxRenderer::render(const Document& doc, const std::string& outputPath) {
  try {
    Presentation prs;

    // 遍历文档中的每一页
    for (size_t i = 0; i < doc.slides.size(); ++i) {
      const auto& slideData = doc.slides[i];

	 		if (i > 0) {
 		  	// 直接从第一张幻灯片获取它正在使用的版式
    		auto* layout = prs.slides()[0].layout_slide(); 
    		if (layout) {
        	prs.slides().add_empty_slide(layout);
   			} else {
        	std::cerr << "[Renderer] 警告：无法获取版式，跳过新建页面" << std::endl;
    		}
			}

      auto& slide = prs.slides()[i];

      // 标题文本框（位置：X=50，Y=50，宽=600，高=80）
      auto& titleShape = slide.shapes().add_auto_shape(
          ShapeType::RECTANGLE, 50, 50, 600, 80);
      titleShape.text_frame()->set_text(slideData.title);

      // 要点文本框（位置：X=50，Y=150，宽=600，高=400）
      if (!slideData.bullets.empty()) {
        std::string bodyText;
        for (size_t j = 0; j < slideData.bullets.size(); ++j) {
          if (j > 0) bodyText += "\n";
          bodyText += "• " + slideData.bullets[j];
        }
        auto& bodyShape = slide.shapes().add_auto_shape( 
            ShapeType::RECTANGLE, 50, 150, 600, 400);
        bodyShape.text_frame()->set_text(bodyText);
      }
    }

    prs.save(outputPath.c_str(), SaveFormat::PPTX);
    std::cout << "[Renderer] 已生成: " << outputPath
              << "，共" << doc.slides.size() << "页\n";
    return true;
    
  } catch (const std::exception& e) {
    std::cerr << "[Renderer] 错误: " << e.what() << "\n";
    return false;
  }
}

}

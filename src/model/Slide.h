#pragma
#include <string> 
#include <vector>

namespace ppt {
struct Slide {
  std::string title;
  std::vector<std::string> bullets;
  
  // 幻灯片布局
  std::string layout = "title_bullets";

  // 每一页结论
  std::string conclusion;
};

} // namespace ppt

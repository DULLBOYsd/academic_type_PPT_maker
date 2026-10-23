#pragma once 
#include "Slide.h"
#include <vector> 

namespace ppt {

struct Document {
  std::string title;
  std::vector<Slide> slides;
};

} // namespace ppt

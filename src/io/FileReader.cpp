#include "FileReader.h"
#include <fstream>
#include <sstream>
#include <stdexcept>

namespace ppt {
  std::string readFile(const std::string& path) {
    std::ifstream ifs(path, std::ios::binary);
    if (!ifs) throw std::runtime_error("无法打开文件夹: " + path);
    std::stringstream ss;
    ss << ifs.rdbuf(); // rdbuf返回当前流对象底层的缓冲区指针
    return ss.str(); // ss.str()获取或设置字符串流内部的字符串内容
  }
}

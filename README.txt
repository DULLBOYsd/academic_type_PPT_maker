调用方式（按顺序执行）
python python\ai_writer.py
python python\image_fetcher.py
python python\render_ppt.py

整体架构

输入文本 (.md / .txt / .docx)
        ↓
┌─────────────────────┐
│   IO 层             │  读取文件，统一成 std::string
├─────────────────────┤
│   Parser 层         │  文本 → std::vector<Slide>
├─────────────────────┤
│   Model 层          │  Slide / Document 数据结构
├─────────────────────┤
│   Renderer 层       │  Slide → .pptx（ZIP + XML）
├─────────────────────┤
│   IO 层             │  写出 .pptx 文件
└─────────────────────┘
        ↓
输出 .pptx

目录结构

text-to-ppt/
├── CMakeLists.txt
├── cmake/
│   └── Dependencies.cmake      # 依赖库配置
├── src/
│   ├── main.cpp
│   ├── model/
│   │   ├── Slide.h             # Slide 结构体
│   │   └── Document.h          # Document 容器
│   ├── parser/
│   │   ├── MarkdownParser.h
│   │   └── MarkdownParser.cpp
│   ├── render/
│   │   ├── PptxRenderer.h
│   │   └── PptxRenderer.cpp
│   └── io/
│     	   ├── FileReader.h
│     	   ├── FileReader.cpp
│   	   ├── FileWriter.h
│    	   └── FileWriter.cpp
├── python/                 # 👈 新增
│   ├── ai_writer.py        # 调用大模型优化大纲
│   ├── render_ppt.py       # 用 python-pptx 渲染最终 PPT
│   └── requirements.txt    # Python 依赖
├── third_party/
│   ├── Aspose.Slides-FOSS-for-Cpp-26.9.0/  	# Slide库依赖
│   ├── pugixml-1.16/                                        	# XML 生成
│   ├── miniz-3.1.2/ 						# 高性能、无损数据压缩库，作为zlib的轻量级替代品
│   └── nlohmann/                    				# C++中JSON提供的简单集成库（只有一个头文件）
│	      └── json.hpp
├── input/
│	└── sample.md
├── output.json                                     		# C++输出中间文件(JSON)
├── tests/
│    	└── test_parser.cpp
└── build/                      # CMake 生成

10.14
目前的一些想法：
目标是实现同根目录下的那个need.png的形式，因为目前这种的看上去最符合网友说的涵盖组会、学术报告、答辩、企业交流，而且并不花哨
1.如何处理模板：可以内嵌一些既有的模板库（比如need.png那种紫色的形式），然后将这些元素首先写进PPT里面，注意放置的位置
			    模板这里可以要求用户提供一个可用的模板（一页就行，然后解析其中的元素、配色等进行复用，然后输出一个类似风格的）
			    如果没有提供就用默认的输出方式
2.如何处理文本：提示词：按照input/sample.md那个格式帮我生成为**主题的.md文件，而且使用*或者其他的.md格式高亮出其中重点的内容，并对.json格式包含的可用形式（如title等）进行拓展
3.文本布局：也可以内置一些应该有的布局形式：比如cover、quote等等
4.公式支持：如果是比较专业的内容，需要考虑latex格式的引入以及公式的推导布局等等，这些是后话


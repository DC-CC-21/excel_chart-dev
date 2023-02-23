import setuptools
  
with open("README.md", "r") as fh:
    description = fh.read()
  
setuptools.setup(
    name="excel_chart",
    version="0.0.1",
    author="DC-CC-21",
    author_email="handstands1000@gmail.com",
    packages=["excel_chart"],
    description="A package for creating excel charts",
    long_description=description,
    long_description_content_type="text/markdown",
    url="https://github.com/DC-CC-21/excel_chart-dev/excel_chart",
    license='MIT',
    python_requires='>=3.8',
    install_requires=["pandas>=1.4.4", "numpy>=1.21.5", "pptx>=0.6.21"]
)
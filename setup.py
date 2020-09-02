import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()
setuptools.setup(
     name='window_input',
     version='0.5',
     author="Kyle Oliver",
     author_email="56kyleoliver@gmail.com",
     description="A set of tools for interacting with windows regardless of them being minimized or in background.",
     long_description=long_description,
     long_description_content_type="text/markdown",
     url="https://github.com/56kyle/window_input",
     packages=['window_input'],
     install_requires=['pywin32'],
     classifiers=[
         "Programming Language :: Python :: 3",
         "License :: OSI Approved :: MIT License",
     ],
 )

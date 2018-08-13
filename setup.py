from setuptools import setup, find_packages


def readme():
    with open('README.md') as f:
        return f.read()


setup(name='cellbase',
      version='0.1.2',
      description='Abstraction layer for accessing spreadsheet as database',
      long_description=readme(),
      long_description_content_type='text/markdown',
      author='imjp94',
      author_email='imjp0921@gmail.com',
      url='https://github.com/imjp94/cellbase',
      license='MIT',
      packages=find_packages(exclude=['tests']),
      install_requires=['openpyxl'],
      zip_safe=False,
      include_package_data=True,
      keywords='spreadsheet excel database query abstraction utility',
      classifiers=[
          'Development Status :: 3 - Alpha',
          'Topic :: Office/Business :: Financial :: Spreadsheet',
          'Topic :: Database',
          'Topic :: Utilities',
          'License :: OSI Approved :: MIT License'
      ])

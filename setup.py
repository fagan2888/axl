from setuptools import setup
import versioneer

setup(name='axl',
      version=versioneer.get_version(),
      cmdclass=versioneer.get_cmdclass(),
      author='Michael C. Grant',
      author_email='mgrant@anaconda.com',
      packages=['axl'],
      license='BSD 3-Clause',
      description='Anaconda for Excel',
      zip_safe=False)

from setuptools import setup

def readme():
    with open('README.rst') as f:
        return f.read()

setup(name='google_dev_utils',
      version='0.9',
      description='Utilities to work with Google services',
      long_description=readme(),
      url='',
      author='Elemental Reasoning',
      author_email='jason@elementalreasoning.com',
      license='MIT',
      packages=['google_dev_utils'],
      install_requires=[
          'google-api-python-client',
          'google-auth-oauthlib'
      ],
      include_package_data=True,
      zip_safe=False)
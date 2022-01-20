from setuptools import setup


setup(
    # Needed to silence warnings (and to be a worthwhile package)
    name='finaigle',
    url='https://github.com/mtlr05/finaigle',
    author='David May',
    author_email='davidmay44@hotmail.com',
    packages=['finaigle'],
    # Needed for dependencies
    install_requires=['numpy','pandas'],
    # *strongly* suggested for sharing
    version='0.2',
    # The license can be anything you like
    license='MIT',
    description='...',
)
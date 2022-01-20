from setuptools import setup, find_packages


setup(
    name='finagle',
    url='https://github.com/mtlr05/finagle',
    packages=find_packages(include=['finagle', 'finagle.*']),
    install_requires=['numpy','pandas'],
    description='for the valuation of a company',
)
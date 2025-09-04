from setuptools import setup, find_packages

setup(
    name='indico-plugin-exportdocs',
    version='0.1.0',
    description='Indico plugin for exporting lists and reports as docx',
    author='Your Name',
    author_email='your@email.com',
    url='https://your-repo-url',
    packages=find_packages(),
    install_requires=[
        'indico',
        'python-docx',
        'docxtpl',
    ],
    entry_points={
        'indico.plugins': [
            'exportdocs = indico_exportdocs.plugin:ExportDocsPlugin',
        ],
    },
    include_package_data=True,
    zip_safe=False,
    license='MIT',
) 
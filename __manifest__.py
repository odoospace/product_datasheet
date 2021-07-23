{
    'name': "Product Datasheet",

    'summary': """
        Module that add extra modification to manage product datasheets
        """,

    'description': """
        *   add extra modification to manage product datasheets

    """,

    'author': "Impulso Diagonal",
    'website': "https://impulso.xyz",

    # for the full list
    'category': 'Extra Tools',
    'version': '13.0.1.35',

    # any module necessary for this one to work correctly
    'depends': ['base', 'sale'],

    # always loaded
    'data': [
        'security/product_datasheet_security.xml',
        'security/ir.model.access.csv',
        'views/templates.xml',
        'views/views.xml',
    ],
}

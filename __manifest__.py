{
    'name': "Heura - Product Tecnical Info",

    'summary': """
        Module that add extra modification to manage Tecnical Info 
        """,

    'description': """
        *   add extra modification to manage Tecnical Info

    """,

    'author': "Impulzia",
    'website': "http://impulzia.com",

    # for the full list
    'category': 'Uncategorized',
    'version': '13.0.1.0.1',

    # any module necessary for this one to work correctly
    'depends': ['base', 'account', 'sale'],

    # always loaded
    'data': [
        # 'security/ir.model.access.csv',
        'views/templates.xml',
    ],
}


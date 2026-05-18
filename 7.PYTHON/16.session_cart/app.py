from flask import Flask, render_template, session, redirect, url_for

app = Flask(__name__)
app.secret_key = 'hello1234'

items = [
    {'id': 'item1', 'name': '햄버거', 'price': 3000},
    {'id': 'item2', 'name': '치즈버거', 'price': 3500},
    {'id': 'item3', 'name': '감자튀김', 'price': 2000},
    {'id': 'item4', 'name': '콜라', 'price': 1500},
    {'id': 'item5', 'name': '치킨너겟', 'price': 4000},
    {'id': 'item6', 'name': '아이스크림', 'price': 2500},
    {'id': 'item7', 'name': '핫도그', 'price': 3200},
    {'id': 'item8', 'name': '피자', 'price': 8000},
    {'id': 'item9', 'name': '스파게티', 'price': 7000},
    {'id': 'item10', 'name': '샐러드', 'price': 4500}
]

@app.route('/')
def index():
    return render_template('product.html', items=items)




@app.route('/add_to_cart/<item_id>')
def add_to_cart(item_id):
    print("장바구니에 담을 상품 :", item_id)
    if 'cart' not in session:
        session['cart'] = {}

    if item_id in session['cart']:

        session['cart'][item_id] += 1
    else:

        session['cart'][item_id] = 1

    print(session['cart'])
    session.modified = True


    return redirect(url_for('index'))


@app.route('/clear_cart')
def clear_cart():

    session.pop('cart', None)

    return redirect(url_for('view_cart'))



@app.route('/cart')


def view_cart():

    cart_items = {}
    total_price = 0

    for item_id, quantity in session.get('cart', {}).items():
        item = next((i for i in items if i['id'] == item_id), None)
        cart_items[item_id] = {
            'name': item['name'],
            'quantity' : quantity,
            'price' : item['price']
        }
        total_price += item['price'] * quantity

    return render_template('cart.html' , cart_items = cart_items, total_price=total_price)


if __name__ == '__main__':
    app.run(debug=True)
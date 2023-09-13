import re

# 示例商品链接
product_url = "https://item.taobao.com/item.htm?id=697975432018&ns=1&abbucket=9#detail"

# 使用正则表达式提取商品ID
match = re.search(r"id=(\d+)", product_url)
if match:
    product_id = match.group(1)
    print("提取到的商品ID:", product_id)
else:
    print("未能提取到商品ID")

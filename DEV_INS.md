# TASK 2 
## BUDGET 100 USD
## Modify the GET_DATA function so it accepts an extra argument called node_identifier
This should be the final version of the GET_DATA function

## Modify the code so the endpoint for the production manifest is https://main-sequence.app as the root

## Implement curve inflation

1. in the sheet called QuantLibBond use the GET_DATA function to  get that data only for 1 day same start same end
2. you will receive something data like this
```shell

curves .iloc[-1:].to_dict()
Out[14]: 
{'time_index': {2727: '2025-12-09T00:00:00Z'},
 'unique_identifier': {2727: 'BANXICO_M_BONOS_OTR'},
 'curve': {2727: 'H4sIAOW5B2kC/zVQSY4gMQj7S86lEpjF0F9rzd/HKalziWIcL/wePz98mft3ZjyeA144Z6LKkhlkP2ftojXcROY4Rdjn6L54A90eDUuLfE7g0x6brDSgosYgPPHROYngRNO3JFOQzLxhVdPJJY2Qay8uXheSN6gsxecQX5zFoLxnlaiMN+XHTy8Wm6YCdvXd/DOgqFYhfqNDH9wjzs++prPbGJfcVfIOuwP6yNIDVO/OO8i8gyZLErFhrc3dVZDfj8Fl0lzrCA0QV8rtzRvGoRrrdGnJTu7urwlpKDHFpez10AaBF6sySuRsi8BXZRSg8GrhbrlBKET8+w95sOsX0AEAAA=='}}


```
3. implement a function that performs the following on a cell INFLATE_ZERO_CURVE
```python
import base64
import gzip
import json
from typing import Any, TypedDict
def decompress_string_to_curve(b64_string: str) -> dict[Any, Any]:
    """
    Decodes, decompresses, and deserializes a string back into a curve dictionary.

    Pipeline: Base64 (text) -> Gzip (binary) -> JSON -> Dict

    Args:
        b64_string: The Base64-encoded string from the database or API.

    Returns:
        The reconstructed Python dictionary.
    """
    # 1. Encode the ASCII string back into Base64 bytes
    base64_bytes = b64_string.encode("ascii")

    # 2. Decode the Base64 to get the compressed Gzip bytes
    compressed_bytes = base64.b64decode(base64_bytes)

    # 3. Decompress the Gzip bytes to get the original JSON bytes
    json_bytes = gzip.decompress(compressed_bytes)

    # 4. Decode the JSON bytes to a string and parse back into a dictionary
    return json.loads(json_bytes.decode("utf-8"))
```

4.This will give you a json with days_to_maturity and curve_value
```python

{'1': 7.749999999998813,
 '27': 7.48835504743776,
 '90': 7.587942448174889,
 '174': 7.62266136204034,
 '321': 7.808454022535802,
 '342': 7.678423783671959,
 '524': 8.305586479770726,
 '692': 8.547975872777657,
 '720': 7.928251689448507,
 '902': 8.415757670504059,
 '1014': 8.789405394462637,
 '1133': 9.000099628116807,
 '1630': 9.718765132742464,
 '1644': 9.677563739306988,
 '1777': 9.782246470162038,
 '2330': 10.440591272791714,
 '2813': 11.079162767733077,
 '6271': 22.294158071760332,
 '10184': 52.61310493723933}
```
so the function should transform each key,value into 2 cells with cell headers days_to_maturity, interest_rate

## Task 3 BUDGET 50 USD

build GET_ASSET()

## Inflate Asset

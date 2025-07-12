from flask import Flask, request, jsonify
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import jieba
from gensim.models import KeyedVectors

app = Flask(__name__)

# 假设模型已加载
file = 'E:\PycharmProjects\\for_VR\\tencent-ailab-embedding-zh-d200-v0.2.0\\tencent-ailab-embedding-zh-d200-v0.2.0/tencent-ailab-embedding-zh-d200-v0.2.0.txt'
kv_model = KeyedVectors.load_word2vec_format(file, binary=False)

@app.route('/vectorize_text', methods=['POST'])
def vectorize_text():
    data = request.get_json()
    text = data.get('text', '')
    if text in kv_model.key_to_index:
        vector = kv_model[text].reshape(1, -1).tolist()
        return jsonify({'vector': vector})
    else:
        words = list(jieba.cut(text))
        word_vectors = [kv_model[word] for word in words if word in kv_model.key_to_index]
        if len(word_vectors) > 0:
            mean_vector = np.mean(word_vectors, axis=0).tolist()
            return jsonify({'vector': mean_vector})
        else:
            return jsonify({'error': 'No valid words found for vectorization'}), 404

@app.route('/similarity', methods=['POST'])
def calculate_similarity():
    data = request.get_json()
    vec1 = np.array(data.get('vec1'))
    vec2 = np.array(data.get('vec2'))
    if vec1 is not None and vec2 is not None:
        similarity = 1 - cosine_similarity(vec1.reshape(1, -1), vec2.reshape(1, -1))[0][0]
        return jsonify({'similarity': similarity})
    else:
        return jsonify({'error': 'Invalid vectors'}), 400

if __name__ == '__main__':
    app.run(port=5000)
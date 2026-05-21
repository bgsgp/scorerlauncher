using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.ML.OnnxRuntime;
using Microsoft.ML.OnnxRuntime.Tensors;
using Newtonsoft.Json;

namespace scorerlauncher
{
    public class NameValidator : IDisposable
    {
        private readonly InferenceSession _session;
        private readonly Dictionary<string, int> _vocab;
        private readonly int _maxLength;
        private readonly string _inputName;
        private readonly string _outputName;
        private bool _disposed;

        public NameValidator(string modelPath, string vocabPath, int maxLength = 8,
            string inputName = "input", string outputName = "output")
        {
            _maxLength = maxLength;
            _inputName = inputName;
            _outputName = outputName;
            _session = new InferenceSession(modelPath);
            _vocab = LoadVocab(vocabPath);
        }

        /// <summary>
        /// 加载词表，支持数组或字典格式，保留原始 token 字符串作为键
        /// </summary>
        private Dictionary<string, int> LoadVocab(string path)
        {
            string json = File.ReadAllText(path);
            if (json.TrimStart().StartsWith('['))
            {
                var list = JsonConvert.DeserializeObject<List<string>>(json);
                if (list == null)
                    throw new InvalidDataException("词表 JSON 数组解析失败。");
                var dict = new Dictionary<string, int>();
                for (int i = 0; i < list.Count; i++)
                    dict[list[i]] = i;
                return dict;
            }
            else
            {
                var dict = JsonConvert.DeserializeObject<Dictionary<string, int>>(json);
                if (dict == null)
                    throw new InvalidDataException("词表 JSON 字典解析失败。");
                return dict;
            }
        }

        /// <summary>将姓名编码为 long 数组，未知字符使用 UNK token</summary>
        private long[] EncodeName(string name)
        {
            int padIdx = _vocab.TryGetValue("<pad>", out int p) ? p : 0;
            int unkIdx = _vocab.TryGetValue("<unk>", out int u) ? u :
                         _vocab.TryGetValue("[UNK]", out u) ? u : padIdx;

            long[] indices = new long[_maxLength];
            for (int i = 0; i < _maxLength; i++)
                indices[i] = padIdx;

            if (string.IsNullOrWhiteSpace(name))
                return indices;

            var trimmed = name.Length > _maxLength ? name.Substring(0, _maxLength) : name;
            for (int i = 0; i < trimmed.Length; i++)
            {
                string token = trimmed[i].ToString();
                indices[i] = _vocab.TryGetValue(token, out int idx) ? idx : unkIdx;
            }
            return indices;
        }

        /// <summary>获取姓名有效的概率（经过 sigmoid 归一化）</summary>
        public float GetProbability(string name)
        {
            long[] indices = EncodeName(name);
            var inputTensor = new DenseTensor<long>(indices, new[] { 1, _maxLength });
            var inputs = new List<NamedOnnxValue>
            {
                NamedOnnxValue.CreateFromTensor(_inputName, inputTensor)
            };
            using var results = _session.Run(inputs);
            var outputTensor = results.First(x => x.Name == _outputName).AsTensor<float>();
            float logit = outputTensor[0, 1];   // 取正类分数，与 Python 的 outputs[0][0,1] 一致
            float prob = 1.0f / (1.0f + MathF.Exp(-logit));
            return prob;
        }

        /// <summary>判断是否为有效姓名（概率 >= 0.5）</summary>
        public bool IsValid(string name)
        {
            return GetProbability(name) >= 0.5f;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                _session?.Dispose();
            }

            _disposed = true;
        }
    }
}
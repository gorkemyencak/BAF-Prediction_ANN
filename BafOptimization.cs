using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeOpenXml;

namespace BafOptimization
{
    public class Optimize
    {
        private int[] layers; //layers    
        private float[][] neurons; //neurons    
        // private float[][] biases; //biasses    
        private float[][][] weights; //weights    
        private int[] activations; //layers

        private List<DataModel> data;
        private DtoModel[] dtos;
        private ConfigFile config;
        private MinMaxScaler[] _scalers;

        private List<DataModel> ReadData()
        {
            using var package = new ExcelPackage(new FileInfo("BAF_training_data_main.xlsx"));

            var firstSheet = package.Workbook.Worksheets["Sheet1"];

            var obj = firstSheet.ConvertSheetToObjects<DataModel>();

            // TODO test et ???
            obj = obj.Select(x =>
            {
                x.Kalinlik = Math.Round(x.Kalinlik, 2);
                return x;
            });

            return obj.ToList();
        }

        private ConfigFile ReadConfig()
        {
            var jsonConfig = File.ReadAllText("config.json");
            return JsonSerializer.Deserialize<ConfigFile>(jsonConfig);
        }

        private MinMaxScaler[] ConstructScaler()
        {
            return config.Scaler.Select((t, i) => new MinMaxScaler(config.Columns[i], t[0], t[1])).ToArray();
        }
        
        public DtoModel[] ConstructDtoModels()
        {
            var list = new List<DtoModel>();
            foreach (var model in data)
            {
                var dto = new DtoModel();
                dto.Genislik = model.Genislik;
                dto.Kalinlik = (float)model.Kalinlik;
                dto.BobinNo = model.BobinNo;
                dto.BobinTonaj = model.BobinTonaj;
                dto.KaideTonaj = model.KaideTonaj;

                dto.ProgramNo = new Dictionary<string, int>();
                foreach (var s in config.Columns.Where(x=>x.StartsWith("ProgNo_")))
                {
                    var value = s[7..] == model.ProgramNo.ToString() ? 1 : 0;
                    dto.ProgramNo.Add(s, value);
                }
                
                dto.KaideSiraNo = new Dictionary<string, int>();
                foreach (var s in config.Columns.Where(x=>x.StartsWith("KaideSiraNo_")))
                {
                    var value = s[12..] == model.KaideSiraNo.ToString() ? 1 : 0;
                    dto.KaideSiraNo.Add(s, value);
                }
                
                list.Add(dto);
            }

            return list.ToArray();
        }

        public DtoModel Convert(DataModel model)
        {
            var returner = new DtoModel
            {
                BobinTonaj = _scalers[0].Fit(model.BobinTonaj),
                KaideTonaj = _scalers[1].Fit(model.KaideTonaj),
                Kalinlik = _scalers[2].Fit((float)model.Kalinlik),
                Genislik = _scalers[3].Fit(model.Genislik),
                BobinNo = model.BobinNo
            };

            returner.ProgramNo = new Dictionary<string, int>();
            foreach (var s in config.Columns.Where(x=>x.StartsWith("ProgNo_")))
            {
                var value = s[7..] == model.ProgramNo.ToString() ? 1 : 0;
                returner.ProgramNo.Add(s, value);
            }
                
            returner.KaideSiraNo = new Dictionary<string, int>();
            foreach (var s in config.Columns.Where(x=>x.StartsWith("KaideSiraNo_")))
            {
                var value = s[12..] == model.KaideSiraNo.ToString() ? 1 : 0;
                returner.KaideSiraNo.Add(s, value);
            }

            return returner;
        }

        public Optimize()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            this.data = ReadData();
            this.config = ReadConfig();
            this.dtos = ConstructDtoModels();
            this._scalers = ConstructScaler();

            this.layers = new int[config.LayerNumber + 1];
            // for (int i = 0; i < layers.Length; i++)
            // {
            //     this.layers[i] = layers[i];
            // }
            

            InitNeurons();
            // InitBiases();
            InitWeights();
        }

        //create empty storage array for the neurons in the network.
        private void InitNeurons()
        {
            List<float[]> neuronsList = new List<float[]>();
            
            // Input Layer
            neuronsList.Add(new float[config.InputLayer]); // feature count
            
            // Hidden layers
            for (int i = 1; i < layers.Length-1; i++)
            {
                neuronsList.Add(new float[config.NodeNumber]);
            }
            
            // Output layer
            neuronsList.Add(new float[1]);

            neurons = neuronsList.ToArray();
        }

        //initializes and populates array for the biases being held within the network.
        // private void InitBiases()
        // {
        //     List<float[]> biasList = new List<float[]>();
        //     for (int i = 0; i < layers.Length; i++)
        //     {
        //         float[] bias = new float[layers[i]];
        //         for (int j = 0; j < layers[i]; j++)
        //         {
        //             bias[j] = UnityEngine.Random.Range(-0.5f, 0.5f);
        //         }
        //
        //         biasList.Add(bias);
        //     }
        //
        //     biases = biasList.ToArray();
        // }

        //initializes random array for the weights being held in the network.
        private void InitWeights()
        {
            // List<float[][]> weightsList = new List<float[][]>();
            // for (int i = 1; i < layers.Length; i++)
            // {
            //     List<float[]> layerWeightsList = new List<float[]>();
            //     int neuronsInPreviousLayer = layers[i - 1];
            //     for (int j = 0; j < neurons[i].Length; j++)
            //     {
            //         float[] neuronWeights = new float[neuronsInPreviousLayer];
            //         for (int k = 0; k < neuronsInPreviousLayer; k++)
            //         {
            //             neuronWeights[k] = (float)new Random().NextDouble();
            //         }
            //
            //         layerWeightsList.Add(neuronWeights);
            //     }
            //
            //     weightsList.Add(layerWeightsList.ToArray());
            // }

            weights = config.Network;
            // weights = weightsList.ToArray();
        }

        public float activate(float value)
        {
            return (float) Math.Max(0, value);
        }

        //feed forward, inputs >==> outputs.
        public float[] FeedForward(float[] inputs)
        {
            for (int i = 0; i < inputs.Length; i++)
            {
                neurons[0][i] = inputs[i];
            }

            // for (int i = 1; i < layers.Length; i++)
            // {
            //     // int layer = i - 1;
            //     for (int j = 0; j < neurons[i].Length; j++)
            //     {
            //         float value = 0f;
            //         for (int k = 0; k < neurons[i - 1].Length; k++)
            //         {
            //             value += weights[i - 1][j][k] * neurons[i - 1][k];
            //         }
            //
            //         neurons[i][j] = activate(value);
            //         // neurons[i][j] = activate(value + biases[i][j]);
            //     }
            // }
            for (int i = 0; i < layers.Length - 1; i++)
            {
                // int layer = i - 1;
                for (int k = 0; k < neurons[i+1].Length; k++)
                {
                    float value = 0f;
                    for (int j = 0; j < neurons[i].Length; j++)
                    {
                        value += weights[i][j][k] * neurons[i][j];
                    }
            
                    neurons[i+1][k] = activate(value);
                    // neurons[i][j] = activate(value + biases[i][j]);
                }
            }

            return neurons[^1];
        }

        public void ConstructModel()
        {

        }

        public double Run()
        {
            var testData = new DataModel()
            {
                /*
                KaideSiraNo = 1,
                ProgramNo = 21,
                Genislik = 1280,
                Kalinlik = 0.32,
                KaideTonaj = 64375,
                BobinTonaj = 24485,
                BobinNo = "A039538A"
                */
                KaideSiraNo = 2, 
                ProgramNo = 58,
                Genislik = 1519,
                Kalinlik = 1.11,
                KaideTonaj = 74110,
                BobinTonaj = 26225,
                BobinNo = "A043997"

            };

            // TODO
            // Construct avaliable programno and kaidesirano fields
            // then query it via incoming request
            // if avaliable data doesnt contain that values throw new argument exception

            var query = Convert(testData);
            
            var inputLayer = new List<float>
            {
                query.BobinTonaj,
                query.KaideTonaj,
                (float)query.Kalinlik,
                query.Genislik,
            };
            inputLayer.AddRange(query.ProgramNo.Select(programNos => programNos.Value).Select(dummy => (float) dummy));
            inputLayer.AddRange(query.KaideSiraNo.Select(kaideSiraNos => kaideSiraNos.Value).Select(dummy => (float) dummy));
            
            var result = FeedForward(inputLayer.ToArray());
            return System.Convert.ToDouble(result[0]) * config.MaxTavTime;
        }
    }

    public class MinMaxScaler
    {
        public string Name;
        private readonly float _max;
        private readonly float _min;
        private float Range => _max - _min;
        
        public MinMaxScaler(float min, float max)
        {
            _max = max;
            _min = min;
        }

        public MinMaxScaler(string name, float min, float max)
        {
            Name = name;
            _max = max;
            _min = min;
        }

        public float Fit(float value)
        {
            return (value - _min) / Range;
        }
    }
    
    public class DataModel
    {
        [Column(2)] public int BobinTonaj { get; set; }
        [Column(3)] public int KaideTonaj { get; set; }
        [Column(4)] public double Kalinlik { get; set; }
        [Column(5)] public int Genislik { get; set; }
        [Column(6)] public int ProgramNo { get; set; }

        [Column(1)] public string BobinNo { get; set; }
        [Column(7)] public int KaideSiraNo { get; set; }
    }

    public class DtoModel
    {
        public float BobinTonaj { get; set; }
        public float KaideTonaj { get; set; }
        public float Kalinlik { get; set; }
        public float Genislik { get; set; }
        // public int ProgramNo { get; set; }
        public Dictionary<string, int> ProgramNo { get; set; }
        public string BobinNo { get; set; }
        // public int KaideSiraNo { get; set; }
        public Dictionary<string, int> KaideSiraNo { get; set; }
    }

    public static class EPPLusExtensions
    {
        public static IEnumerable<T> ConvertSheetToObjects<T>(this ExcelWorksheet worksheet) where T : new()
        {
            Func<CustomAttributeData, bool> columnOnly = y => y.AttributeType == typeof(Column);

            var columns = typeof(T)
                .GetProperties()
                .Where(x => x.CustomAttributes.Any(columnOnly))
                .Select(p => new
                {
                    Property = p,
                    Column = p.GetCustomAttributes<Column>().First().ColumnIndex //safe because if where above
                }).ToList();


            var rows = worksheet.Cells
                .Select(cell => cell.Start.Row)
                .Distinct()
                .OrderBy(x => x);


            //Create the collection container
            var collection = rows.Skip(1)
                .Select(row =>
                {
                    var tnew = new T();
                    columns.ForEach(col =>
                    {
                        //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                        var val = worksheet.Cells[row, col.Column];
                        //If it is numeric it is a double since that is how excel stores all numbers
                        if (val.Value == null)
                        {
                            col.Property.SetValue(tnew, null);
                            return;
                        }

                        if (col.Property.PropertyType == typeof(Int32))
                        {
                            col.Property.SetValue(tnew, val.GetValue<int>());
                            return;
                        }

                        if (col.Property.PropertyType == typeof(double))
                        {
                            col.Property.SetValue(tnew, val.GetValue<double>());
                            return;
                        }

                        if (col.Property.PropertyType == typeof(DateTime))
                        {
                            col.Property.SetValue(tnew, val.GetValue<DateTime>());
                            return;
                        }

                        //Its a string
                        col.Property.SetValue(tnew, val.GetValue<string>());
                    });

                    return tnew;
                });


            //Send it back
            return collection;
        }
    }

    [AttributeUsage(AttributeTargets.All)]
    public class Column : System.Attribute
    {
        public int ColumnIndex { get; set; }


        public Column(int column)
        {
            ColumnIndex = column;
        }
    }
    
    public class ConfigFile
    {
        [JsonPropertyName("NN")]
        public float[][][] Network { get; set; }
        
        [JsonPropertyName("input_layer")]
        public int InputLayer { get; set; }
        
        [JsonPropertyName("layer_number")]
        public int LayerNumber { get; set; }
        
        [JsonPropertyName("node_number")]
        public int NodeNumber { get; set; }
        
        [JsonPropertyName("columns")]
        public string[] Columns { get; set; }
        
        [JsonPropertyName("scaler")]
        public float[][] Scaler { get; set; }
        
        [JsonPropertyName("maxtavtime")]
        public float MaxTavTime { get; set; }
    }
}
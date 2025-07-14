digitDatasetPath = fullfile(toolboxdir('nnet'), 'nndemos', 'nndatasets', 'DigitDataset');
imds = imageDatastore(digitDatasetPath, 'IncludeSubfolders', true, 'LabelSource', 'foldernames');
[imdsTrain, imdsTest] = splitEachLabel(imds, 0.8, 'randomized');

figure;
perm = randperm(10000,25);
for i = 1:25
    subplot(5,5,i);
    imshow(imds.Files{perm(i)});
end

layers = [
    imageInputLayer([28 28 1], 'Name', 'input') % 输入层，图像大小为28x28，灰度图像

    convolution2dLayer(3, 8, 'Padding', 'same', 'Name', 'conv_1') % 卷积层
    batchNormalizationLayer('Name', 'BN_1') % 批量归一化层
    reluLayer('Name', 'relu_1') % ReLU激活层
    maxPooling2dLayer(2, 'Stride', 2, 'Name', 'maxpool_1') % 最大池化层

    convolution2dLayer(3, 16, 'Padding', 'same', 'Name', 'conv_2') % 卷积层
    batchNormalizationLayer('Name', 'BN_2') % 批量归一化层
    reluLayer('Name', 'relu_2') % ReLU激活层
    maxPooling2dLayer(2, 'Stride', 2, 'Name', 'maxpool_2') % 最大池化层

    fullyConnectedLayer(10, 'Name', 'fc') % 全连接层，输出10个类别
    softmaxLayer('Name', 'softmax') % Softmax层
    classificationLayer('Name', 'output') % 分类输出层
];


options = trainingOptions('sgdm', ...
    'MaxEpochs', 10, ...
    'Shuffle', 'every-epoch', ...
    'ValidationData', imdsTest, ...
    'ValidationFrequency', 30, ...
    'Verbose', false, ...
    'Plots', 'training-progress');

net = trainNetwork(imdsTrain, layers, options);

YTest = classify(net, imdsTest);
accuracy = sum(YTest == imdsTest.Labels) / numel(imdsTest.Labels);
fprintf('准确率: %.2f%%\n', accuracy * 100);


figure;
plot(trainingInfo.TrainingLoss);
hold on;
plot(trainingInfo.ValidationLoss);
legend('Training Loss', 'Validation Loss');
xlabel('Iteration');
ylabel('Loss');
title('训练过程中的损失');

figure;
plot(trainingInfo.TrainingAccuracy);
hold on;
plot(trainingInfo.ValidationAccuracy);
legend('Training Accuracy', 'Validation Accuracy');
xlabel('Iteration');
ylabel('Accuracy');
title('训练过程中的准确率');





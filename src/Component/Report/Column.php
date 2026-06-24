<?php

namespace App\Component\Report;

class Column implements \JsonSerializable
{
    public const TYPE_BOOLEAN = 'boolean';
    public const TYPE_DATE = 'date';
    public const TYPE_FLOAT = 'float';
    public const TYPE_INTEGER = 'integer';
    public const TYPE_STRING = 'string';
    public const TYPE_AUTO_INCREMENT = 'auto';
    public const TYPE_GROUP = 'group';

    public const FORMAT_CURRENCY = 'currency';
    public const FORMAT_EMAIL = 'email';
    public const FORMAT_NUMBER = 'number';
    public const FORMAT_URL = 'url';

    protected string $name;
    protected string $display;
    protected string $type;
    protected ?string $format = null;
    protected array $options = [];
    protected bool $total = false;
    protected string $totalType = 'sum';
    protected bool $subTotal = false;
    protected array $subTotalGroup = [];
    protected string $totalLabel = 'Total';
    protected string $subTotalLabel = 'Sub-total';
    protected string $dateInputFormat = 'd-m-Y';
    protected bool $wrapText = false;
    protected bool $isLargeNumber = false;

    public function __construct(string $name, string $display, string $type, array $options = [])
    {
        $this
            ->setName($name)
            ->setDisplay($display)
            ->setType($type)
            ->setOptions($options);
    }

    public function setName(string $name): self
    {
        $this->name = $name;
        return $this;
    }

    public function getName()
    {
        return $this->name;
    }

    public function setDisplay($display)
    {
        $this->display = $display;
        return $this;
    }

    public function getDisplay()
    {
        return $this->display;
    }

    public function setType($type): self
    {
        $types = [
            self::TYPE_BOOLEAN,
            self::TYPE_DATE,
            self::TYPE_FLOAT,
            self::TYPE_INTEGER,
            self::TYPE_STRING,
            self::TYPE_AUTO_INCREMENT,
            self::TYPE_GROUP,
        ];

        if (!in_array($type, $types)) {
            throw new \InvalidArgumentException("Invalid column type: '$type'");
        }

        $this->type = $type;
        return $this;
    }

    public function getType()
    {
        return $this->type;
    }

    public function getFormat()
    {
        return $this->format;
    }

    public function isTotal(): bool
    {
        return $this->total;
    }

    public function isSubTotal(): bool
    {
        return $this->subTotal;
    }

    public function getSubTotalGroup(): array
    {
        return $this->subTotalGroup;
    }

    public function getChoices()
    {
        return $this->options['choices'] ?? [];
    }

    public function jsonSerialize(): array
    {
        return [
            'name' => $this->getName(),
            'display' => $this->getDisplay(),
            'type' => $this->getType(),
            'format' => $this->getFormat(),
            'total' => $this->isTotal(),
            'choices' => $this->getChoices(),
        ];
    }

    public function setOptions(array $options)
    {
        $this->options = $options;
        $this->total = $options['total'] ?? false;
        $this->format = $options['format'] ?? null;
        $this->subTotal = $options['subTotal'] ?? false;
        $this->subTotalGroup = $options['subTotalGroup'] ?? [];
        $this->subTotalLabel = $options['subTotalLabel'] ?? $this->subTotalLabel;
        $this->totalLabel = $options['totalLabel'] ?? $this->totalLabel;
        $this->wrapText = $options['wrapText'] ?? false;
        $this->dateInputFormat = $options['inputFormat'] ?? 'd-m-Y';
        $this->isLargeNumber = $options['isLargeNumber'] ?? false;
        if ($this->total) {
            $this->totalType = $options['totalType'] ?? 'sum';
        }
    }

    public function getSubTotalLabel(): string
    {
        return $this->subTotalLabel;
    }

    public function getTotalLabel(): string
    {
        return $this->totalLabel;
    }

    public function getWidth()
    {
        return $this->options['width'] ?? null;
    }

    public function isWrapText(): bool
    {
        return $this->wrapText;
    }

    public function setWrapText(bool $wrapText): void
    {
        $this->wrapText = $wrapText;
    }

    public function getIsLargeNumber(): bool
    {
        return $this->isLargeNumber;
    }

    public function getInputFormat()
    {
        return $this->dateInputFormat;
    }

    public function isCount(): bool
    {
        return 'count' === $this->totalType;
    }
}
